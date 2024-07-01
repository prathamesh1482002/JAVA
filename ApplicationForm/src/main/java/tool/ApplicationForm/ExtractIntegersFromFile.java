package tool.ApplicationForm;

import java.io.*;
import java.util.regex.*;
import java.util.ArrayList;
import java.util.List;

public class ExtractIntegersFromFile {
    public static void main(String[] args) {
        String inputFilePath = "C:\\File Handling\\Text1.txt";
        String outputFilePath = "C:\\File Handling\\newewejwkjewjeh.txt";

        try {
            // Check if input file exists
            File inputFile = new File(inputFilePath);
            if (!inputFile.exists()) {
                System.out.println("Input file not found.");
                return;
            }

            // Prepare to read the input file
            BufferedReader reader = new BufferedReader(new FileReader(inputFile));
            BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath));

            String line;
            Pattern pattern = Pattern.compile("\\d+");

            // Process each line
            while ((line = reader.readLine()) != null) {
                Matcher matcher = pattern.matcher(line);

                // Find and write each integer to the output file
                while (matcher.find()) {
                    writer.write(matcher.group());
                    writer.newLine();
                }
            }

            // Close readers and writers
            reader.close();
            writer.close();

            System.out.println("Integers extracted and written to output.txt successfully.");
        } catch (IOException e) {
            System.out.println("An error occurred: " + e.getMessage());
        }
    }
}
//