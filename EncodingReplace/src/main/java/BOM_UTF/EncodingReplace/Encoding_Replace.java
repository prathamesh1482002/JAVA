package BOM_UTF.EncodingReplace;
import java.awt.Desktop;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;

public class Encoding_Replace {
	public static void main(String directoryPath) {
	//public static void main(String[] args) {
        //String directoryPath = "D:\\Journals";
        Path path = Paths.get(directoryPath);
        if (!Files.exists(path) || !Files.isDirectory(path)) {
            System.exit(1);
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(path, "*.xml")) {
            for (Path entry : stream) {
                File xmlFile = entry.toFile();
                removeUnwantedData(xmlFile);
                openInNotepad(xmlFile);
            }
        } catch (IOException e) {	
            e.printStackTrace();
        }        
        closeNotepadInstances();
    }
	private static void removeUnwantedData(File file) throws IOException {
        File tempFile = new File(file.getAbsolutePath() + ".tmp");

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file), StandardCharsets.ISO_8859_1));
             BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tempFile), StandardCharsets.ISO_8859_1))) {
            String line = reader.readLine();
            if (line != null) {
                int index = line.indexOf('<');
                if (index != -1) {
                    writer.write(line.substring(index));
                }
            }
            String restOfLine;
            while ((restOfLine = reader.readLine()) != null) {
                writer.write(restOfLine);
                writer.newLine();
            }
        }

        if (!file.delete()) {
            System.err.println("Failed to delete original file: " + file.getAbsolutePath());
            return;
        }

        if (!tempFile.renameTo(file)) {
            System.err.println("Failed to rename temp file: " + tempFile.getAbsolutePath());
        }
    }

    private static void openInNotepad(File file) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop desktop = Desktop.getDesktop();
                if (desktop.isSupported(Desktop.Action.EDIT)) {	
                    desktop.edit(file);
                } else {
                    System.err.println("Not supported on this platform.");
                }
            } else {
                System.err.println("Not supported on this platform.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void closeNotepadInstances() {
        try {
            ProcessBuilder processBuilder = new ProcessBuilder("taskkill", "/F", "/IM", "notepad.exe");
            Process process = processBuilder.start();
            process.waitFor();
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}