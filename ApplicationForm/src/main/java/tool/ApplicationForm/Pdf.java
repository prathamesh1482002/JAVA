package tool.ApplicationForm;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.io.*;

public class Pdf extends JPanel {
	public Pdf() {
	}

    private static final long serialVersionUID = 1L;
    private BufferedImage pageImage;
    private BufferedImage scaledPageImage;
    private Rectangle selectionRect;
    private int[] pageStartY;
    private static File pdfFile;
    private double scale = 1.0;
    private static JFrame currentFrame = null;

    public static Pdf createPdfViewer(File file) {
        Pdf pdfViewer = new Pdf();
        pdfViewer.initialize(file);
        return pdfViewer;
    }

    public void initialize(File file) {
        pdfFile = file;
        setLayout(null);
        setBackground(Color.WHITE);

        addMouseListener(new MouseAdapter() {
            public void mousePressed(MouseEvent e) {
                selectionRect = new Rectangle(e.getPoint());
            }

            public void mouseReleased(MouseEvent e) {
                if (selectionRect != null) {
                    copySelectedText();
                    selectionRect = null;
                }
            }
        });

        addMouseMotionListener(new MouseAdapter() {
            public void mouseDragged(MouseEvent e) {
                if (selectionRect != null) {
                    selectionRect.setSize(e.getX() - selectionRect.x, e.getY() - selectionRect.y);
                    repaint();
                }
            }
        });

        loadPdf(file);
    }

    public void loadPdf(File file) {
        try (PDDocument document = PDDocument.load(file)) {
            PDFRenderer renderer = new PDFRenderer(document);
            int numPages = document.getNumberOfPages();
            int width = 0;
            int height = 0;
            pageStartY = new int[numPages];

            for (int i = 0; i < numPages; i++) {
                BufferedImage pageImage = renderer.renderImageWithDPI(i, 72);
                width = Math.max(width, pageImage.getWidth());
                pageStartY[i] = height;
                height += pageImage.getHeight();

                if (i == 0) {
                    this.pageImage = pageImage;
                } else {
                    this.pageImage = joinBufferedImage(this.pageImage, pageImage);
                }
            }
            setPreferredSize(new Dimension((int) (width * scale), (int) (height * scale)));
            revalidate();
            repaint();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private BufferedImage joinBufferedImage(BufferedImage img1, BufferedImage img2) {
        int width = Math.max(img1.getWidth(), img2.getWidth());
        int height = img1.getHeight() + img2.getHeight();
        BufferedImage combined = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);

        Graphics2D g = combined.createGraphics();
        g.drawImage(img1, 0, 0, null);
        g.drawImage(img2, 0, img1.getHeight(), null);
        g.dispose();

        return combined;
    }

    private void copySelectedText() {
        try (PDDocument document = PDDocument.load(pdfFile)) {
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.setSortByPosition(true);
            StringBuilder selectedText = new StringBuilder();
            for (int i = 0; i < document.getNumberOfPages(); i++) {
                PDPage page = document.getPage(i);
                int pageHeight = (int) page.getMediaBox().getHeight();
                int x1 = (int) (selectionRect.x / scale);
                int y1 = (int) ((selectionRect.y - pageStartY[i] * scale) / scale);
                int x2 = (int) (x1 + selectionRect.width / scale);
                int y2 = (int) (y1 + selectionRect.height / scale);
                if (y1 >= 0 && y2 <= pageHeight) {
                    Rectangle rect = new Rectangle(x1, y1, x2 - x1, y2 - y1);
                    stripper.addRegion("selection" + i, rect);
                    stripper.extractRegions(page);
                    selectedText.append(stripper.getTextForRegion("selection" + i));
                }
            }
            if (!selectedText.toString().isEmpty()) {
                StringSelection stringSelection = new StringSelection(selectedText.toString());
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                clipboard.setContents(stringSelection, null);
            } else {
                JOptionPane.showMessageDialog(this, "No text found in this region.",
                        "Error", JOptionPane.ERROR_MESSAGE);
            }
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Failed to copy text: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }


    protected void paintComponent(Graphics g) {
        super.paintComponent(g);
        if (pageImage != null) {
            if (scaledPageImage == null) {
                scaledPageImage = getScaledImage(pageImage, scale);
            }
            g.drawImage(scaledPageImage, 0, 0, this);
            if (selectionRect != null) {
                g.setColor(new Color(0, 0, 255, 64));
                g.fillRect(selectionRect.x, selectionRect.y, selectionRect.width, selectionRect.height);
            }
        }
    }

    private BufferedImage getScaledImage(BufferedImage src, double scale) {
        int w = (int) (src.getWidth() * scale);
        int h = (int) (src.getHeight() * scale);
        BufferedImage scaledImage = new BufferedImage(w, h, src.getType());
        Graphics2D g2d = scaledImage.createGraphics();
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.drawImage(src, 0, 0, w, h, null);
        g2d.dispose();
        return scaledImage;
    }

    private void zoomIn() {
        scale *= 1.25;
        scaledPageImage = getScaledImage(pageImage, scale);
        setPreferredSize(new Dimension((int) (pageImage.getWidth() * scale), (int) (pageImage.getHeight() * scale)));
        revalidate();
        repaint();
    }

    private void zoomOut() {
        scale /= 1.25;
        scaledPageImage = getScaledImage(pageImage, scale);
        setPreferredSize(new Dimension((int) (pageImage.getWidth() * scale), (int) (pageImage.getHeight() * scale)));
        revalidate();
        repaint();
    }

    public static void openFile(File file) {
        if (currentFrame != null) {
            currentFrame.dispose();
        }
        SwingUtilities.invokeLater(() -> {
            currentFrame = new JFrame("PDF View");
            currentFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            currentFrame.setLocation(830, 50);
            Pdf pdfViewer = Pdf.createPdfViewer(file);

            JScrollPane scrollPane = new JScrollPane(pdfViewer);
            scrollPane.getVerticalScrollBar().setUnitIncrement(50);
            scrollPane.getVerticalScrollBar().setBlockIncrement(50);
            scrollPane.setPreferredSize(new Dimension(708, 840));

            JPanel controls = new JPanel();
            JButton zoomInButton = new JButton("Zoom In");
            JButton zoomOutButton = new JButton("Zoom Out");

            zoomInButton.addActionListener(e -> pdfViewer.zoomIn());
            zoomOutButton.addActionListener(e -> pdfViewer.zoomOut());

            controls.add(zoomInButton);
            controls.add(zoomOutButton);

            currentFrame.add(controls, BorderLayout.NORTH);
            currentFrame.add(scrollPane, BorderLayout.CENTER);
            currentFrame.pack();
            currentFrame.setVisible(true);
        });
    }
}