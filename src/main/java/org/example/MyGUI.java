package org.example;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import javax.swing.*;

public class MyGUI extends JFrame {
    private final JLabel inputLabel;
    private final JLabel outputLabel;
    private final JTextField inputFile;
    private final JTextField outputFile;
    private final JButton browseInputButton;
    private final JButton browseOutputButton;
    private final JButton convertButton;

    public MyGUI() {
        // Set up the GUI components
        inputLabel = new JLabel("Input file:");
        outputLabel = new JLabel("Output directory:");
        inputFile = new JTextField();
        outputFile = new JTextField();
        browseInputButton = new JButton("Browse");
        browseOutputButton = new JButton("Browse");
        convertButton = new JButton("Convert");

        // Add the components to the GUI
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(5, 5, 5, 5);
        c.gridx = 0;
        c.gridy = 0;
        panel.add(inputLabel, c);
        c.gridx = 1;
        c.gridy = 0;
        c.weightx = 1;
        c.fill = GridBagConstraints.HORIZONTAL;
        panel.add(inputFile, c);
        c.gridx = 2;
        c.gridy = 0;
        c.weightx = 0;
        c.fill = GridBagConstraints.NONE;
        panel.add(browseInputButton, c);
        c.gridx = 0;
        c.gridy = 1;
        c.weightx = 0;
        c.fill = GridBagConstraints.NONE;
        panel.add(outputLabel, c);
        c.gridx = 1;
        c.gridy = 1;
        c.weightx = 1;
        c.fill = GridBagConstraints.HORIZONTAL;
        panel.add(outputFile, c);
        c.gridx = 2;
        c.gridy = 1;
        c.weightx = 0;
        c.fill = GridBagConstraints.NONE;
        panel.add(browseOutputButton, c);
        c.gridx = 1;
        c.gridy = 2;
        c.weightx = 0;
        c.fill = GridBagConstraints.NONE;
        panel.add(convertButton, c);
        add(panel);

        // Add action listeners to the buttons
        browseInputButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser = new JFileChooser();
                int result = chooser.showOpenDialog(MyGUI.this);
                if (result == JFileChooser.APPROVE_OPTION) {
                    inputFile.setText(chooser.getSelectedFile().getAbsolutePath());
                }
            }
        });
        browseOutputButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int result = chooser.showOpenDialog(MyGUI.this);
                if (result == JFileChooser.APPROVE_OPTION) {
                    outputFile.setText(chooser.getSelectedFile().getAbsolutePath());
                }
            }
        });
        convertButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String inputLocation = inputFile.getText();
                String outputLocation = outputFile.getText() + "/";
                try {
                    if (Main(inputLocation, outputLocation) == 0) {
                        JOptionPane.showMessageDialog(MyGUI.this, "Error converting file", "Error", JOptionPane.ERROR_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(MyGUI.this, "File converted successfully", "Success", JOptionPane.INFORMATION_MESSAGE);
                    }
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }

            }
        });

        // Set the window properties
        setTitle("Excel to PDF Converter");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        pack();
        setLocationRelativeTo(null);
        setVisible(true);
    }


 /*
    private static void createAndShowGUI() {
        // Create the main frame
        JFrame frame = new JFrame("My Swing Application");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Create a label
        JLabel label = new JLabel("Hello, world!");
        frame.getContentPane().add(label);

        // Pack and display the frame
        frame.pack();
        frame.setVisible(true);
    }

  */
    //public static void main(String[] args) {
    //    SwingUtilities.invokeLater(new Runnable() {
    //        public void run() {
    //            MyGUI gui = new MyGUI();
    //        }
    //    });
    //}
    ///////////////////////////////////////////////////////////////////////////////////////////////////////
    public static int PrintFile(String[][] input, String outputLocation) throws IOException
    {
        String Date = "";
        //need to send data to an array to store all header information for printing
        int numCols = input[0].length;
        String[] header = new String[numCols];

        // Create a PDF document for each row
        for (int i = 0; i < input.length; i++) {
            // Create a new PDF document
            PDDocument document = new PDDocument();

            // Create a new page
            PDPage page = new PDPage();
            document.addPage(page);

            // Create a new content stream for the page
            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            // Set font and font size
            contentStream.setFont(PDType1Font.HELVETICA_BOLD_OBLIQUE, 12);

            File imageFile = new File("Header.png");
            PDImageXObject image = PDImageXObject.createFromFile(imageFile.getAbsolutePath(), document);
            contentStream.drawImage(image, 0, 610, 602,151);

            // Write text to the page
            contentStream.beginText();

            contentStream.newLineAtOffset(100, 600);
            for (int j = 0; j < input[i].length; j++) {

                if (i  == 0 ) header[j] = input[i][j];

                if (i > 0) {
                    if ((j == 10) && !(input[i][j].equals("Degaussing Time"))) {
                        //java.util.Date javaDate = DateUtil.getJavaDate(Double.valueOf(input[i][j]));
                        //Date = new SimpleDateFormat("HH:mm, dd-MMMM-YYYY").format(javaDate);
                        //input[i][j] = Date.toString();
                        //Date = new SimpleDateFormat("dd-MMMM-YYYY").format(javaDate);

                        Date = input[i][j];
                    }

                    String text = input[i][j];
                    if (text != null) {
                        contentStream.showText(header[j]);
                        contentStream.showText(" : ");
                        contentStream.showText(input[i][j]);
                    }
                    contentStream.newLineAtOffset(0, -20);
                }
            }
            contentStream.endText();
            //File imageFileFooter = new File("C:\\Users\\hilton.XPERIEN\\IdeaProjects\\PDFBOX\\src\\Xperien-Logo.png");
            //PDImageXObject imageFooter = PDImageXObject.createFromFile(imageFileFooter.getAbsolutePath(), document);
            //contentStream.drawImage(imageFooter, 50, 630, 502,141);

            //float pageHeight = page.getMediaBox().getHeight();
            float imageHeight = 141; // image height in points

            // calculate the y coordinate to place the image at the bottom of the page
            float y = imageHeight;
            float x = 0;

            File imageFileFooter = new File("Footer.png");
            PDImageXObject imageFooter = PDImageXObject.createFromFile(imageFileFooter.getAbsolutePath(), document);
            contentStream.drawImage(imageFooter, x, y-120, 602,151);


            // Close the content stream
            contentStream.close();

            // Save the document
            if (i > 0) document.save(outputLocation + "DegaussedSN_" + input[i][4] + "_" + Date + "_" + input[i][17]+ ".pdf");
            // Close the document
            document.close();
        }
        return 1;
    }

    public static int Main(String inputLocation, String outputLocation) throws IOException {

        // Open the Excel file
        FileInputStream input = new FileInputStream(new File(inputLocation));
        //XSSFWorkbook workbook = new XSSFWorkbook(input);

        Sheet sheet;
        try (Workbook workbook = WorkbookFactory.create(input)) {

            // Get the first sheet
            sheet = workbook.getSheetAt(0);
        }

        // Define the starting and ending rows to remove
        int startRow = 0;
        int endRow = 1;

        // Remove the first two rows
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
            }
        }

        // Get the number of columns and rows
        int numColumns = 0;
        int numRows = 0;
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                numColumns = Math.max(numColumns, row.getLastCellNum());
                numRows++;
            }
        }

        // Create a 2D array to store the data
        String[][] data = new String[numRows][numColumns];

        // Copy the data from the sheet to the array
        int rowIndex = 0;
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < numColumns; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        data[rowIndex][j] = cell.toString();
                    }
                }
                rowIndex++;
            }
        }

        // Print the data to the File
        if (PrintFile(data, outputLocation) == 1) {
            return numRows;
        } else {
            System.out.println("Print to file not done");
            return 0;
        }
    }
}