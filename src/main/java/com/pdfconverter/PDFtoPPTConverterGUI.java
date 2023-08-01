/*
 Version: 5.0
 Title: Demo version of CP317 project
 Author: @Jacob Schwartz
 Description: This program converts a pdf research report into a powerpoint presentation, 
 We have utilized a graphical user interface to help implement in a user friendly enviorment.
 */
package com.pdfconverter;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.sl.usermodel.PictureData;

public class PDFtoPPTConverterGUI extends JFrame {
    private JPanel cardPanel;
    private CardLayout cardLayout;
    private JButton startButton;

    private JTextField pdfFilePathField;
    private JTextField outputDirField;
    private JButton browsePDFButton;
    private JButton browseOutputButton;
    private JButton convertButton;
    private JTextArea statusTextArea;

    public PDFtoPPTConverterGUI() {
        // Setting the size of the windows 
        setTitle("PDF to PPT Converter Group 8");
        setSize(800, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Create card layout to switch between title and main program
        cardLayout = new CardLayout();
        cardPanel = new JPanel(cardLayout);

        // Add title page to the card layout
        JPanel titlePanel = createTitlePage();
        cardPanel.add(titlePanel, "title");

        // Add main program GUI to the card layout
        JPanel mainPanel = createMainProgramPanel();
        cardPanel.add(mainPanel, "main");

        // Show the title page initially
        cardLayout.show(cardPanel, "title");

        // Add card panel to the main frame
        add(cardPanel, BorderLayout.CENTER);
    }

    private JPanel createTitlePage() {
        JPanel titlePagePanel = new JPanel(new BorderLayout());
        titlePagePanel.setBackground(new Color(41, 128, 185));
        titlePagePanel.setBorder(new EmptyBorder(100, 100, 100, 100));

        JLabel titleLabel = new JLabel("PDF to PPT Converter Group 8");
        titleLabel.setFont(new Font("Arial", Font.BOLD, 36));
        titleLabel.setForeground(Color.WHITE);
        titleLabel.setHorizontalAlignment(SwingConstants.CENTER);
        titlePagePanel.add(titleLabel, BorderLayout.CENTER);

        startButton = new JButton("Start");
        startButton.setFont(new Font("Arial", Font.PLAIN, 18));
        startButton.setForeground(Color.red);
        startButton.setBackground(new Color(39, 174, 96));
        startButton.setBorderPainted(false);
        startButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Show the main program when Start button is clicked
                cardLayout.show(cardPanel, "main");
            }
        });
        JPanel buttonPanel = new JPanel();
        buttonPanel.add(startButton);
        titlePagePanel.add(buttonPanel, BorderLayout.SOUTH);

        return titlePagePanel;
    }

    private JPanel createMainProgramPanel() {
        JPanel mainProgramPanel = new JPanel(new BorderLayout());
        mainProgramPanel.setBackground(new Color(41, 128, 185));
        mainProgramPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        // Create components for the main program GUI
               // Set window properties
        setTitle("PDF To PPT Converter:");
        setSize(800, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Create components
        JLabel pdfFileLabel = new JLabel("PDF File:");
        pdfFileLabel.setFont(new Font("Arial", Font.BOLD, 16));

        JLabel outputDirLabel = new JLabel("Output Directory:");
        outputDirLabel.setFont(new Font("Arial", Font.BOLD, 16));

        pdfFilePathField = new JTextField(20);
        outputDirField = new JTextField(20);

        browsePDFButton = new JButton("Browse Here");
        browsePDFButton.setFont(new Font("Arial", Font.PLAIN, 14));

        browseOutputButton = new JButton("Browse Here");
        browseOutputButton.setFont(new Font("Arial", Font.PLAIN, 14));

        convertButton = new JButton("Please click me to convert");
        convertButton.setFont(new Font("Arial", Font.BOLD, 18));
        convertButton.setForeground(Color.red);
        convertButton.setBackground(new Color(39, 174, 96));
        convertButton.setBorderPainted(false);
        
        statusTextArea = new JTextArea(8, 40);
        statusTextArea.setFont(new Font("Arial", Font.PLAIN, 14));
        statusTextArea.setEditable(false);
        //JScrollPane statusScrollPane = new JScrollPane(statusTextArea);

         // Create panels
        JPanel inputPanel = new JPanel(new GridLayout(2, 3, 10, 10));
        inputPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 20, 10));
        buttonPanel.setBorder(new EmptyBorder(0, 0, 20, 0));

        JPanel statusPanel = new JPanel(new BorderLayout());
        statusPanel.setBorder(new EmptyBorder(0, 20, 20, 20));
        inputPanel.add(pdfFileLabel);
        inputPanel.add(pdfFilePathField);
        inputPanel.add(browsePDFButton);
        inputPanel.add(outputDirLabel);
        inputPanel.add(outputDirField);
        inputPanel.add(browseOutputButton);

        buttonPanel.add(convertButton);

        statusPanel.add(new JScrollPane(statusTextArea), BorderLayout.CENTER);

        // Add panels to main program panel
        mainProgramPanel.add(inputPanel, BorderLayout.NORTH);
        mainProgramPanel.add(buttonPanel, BorderLayout.CENTER);
        mainProgramPanel.add(statusPanel, BorderLayout.SOUTH);

        // Attach event listeners
        browsePDFButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    pdfFilePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
                }
            }
        });

        browseOutputButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    outputDirField.setText(fileChooser.getSelectedFile().getAbsolutePath());
                }
            }
        });

        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String pdfFilePath = pdfFilePathField.getText();
                String outputDir = outputDirField.getText();

                try {
                // Load the PDF document
                PDDocument document = PDDocument.load(new File(pdfFilePath));
                PDFRenderer pdfRenderer = new PDFRenderer(document);

                // Create a PowerPoint presentation
                XMLSlideShow ppt = new XMLSlideShow();

                for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
                    // Convert each PDF page to an image
                    BufferedImage image = pdfRenderer.renderImage(pageIndex);

                    // Get the slide size of PowerPoint (in pixels)
                    int slideWidth = 720; 
                    int slideHeight = 540; 

                    // Create a new BufferedImage with the desired dimensions
                    BufferedImage resizedImage = new BufferedImage(slideWidth, slideHeight, BufferedImage.TYPE_INT_RGB);

                    // Create a graphics context from the resized image
                    Graphics2D g2d = resizedImage.createGraphics();

                    // Draw the original image onto the resized image with the desired dimensions
                    g2d.drawImage(image, 0, 0, slideWidth, slideHeight, null);

                    // Dispose of the graphics context to free resources
                    g2d.dispose();

                    // Create a new slide in the presentation
                    XSLFSlide slide = ppt.createSlide();

                    // Obtain byte array representation
                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    ImageIO.write(resizedImage, "png", baos);
                    byte[] pictureData = baos.toByteArray();
                    baos.close();
                    
                    PictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                    XSLFPictureShape pictureShape = slide.createPicture(pd);
  
                    // Reposition the image at the top-left corner of the slide:
                    pictureShape.setAnchor(new java.awt.Rectangle(0, 0, slideWidth, slideHeight));

                    // Delete the temporary image file
                    File tempImage = new File("temp.png");
                    tempImage.delete();
                }

                // Save the PowerPoint presentation
                FileOutputStream out = new FileOutputStream(outputDir + File.separator + "output_ppt.pptx");
                ppt.write(out);
                out.close();

                // Close the PDF document
                document.close();
                ppt.close();

                statusTextArea.setText("Conversion complete! PowerPoint file saved at: " + outputDir);
            } catch (IOException ex) {
                ex.printStackTrace();
                statusTextArea.setText("Error during conversion: " + ex.getMessage());
            }
        }

        });

        return mainProgramPanel;
    }


    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                PDFtoPPTConverterGUI gui = new PDFtoPPTConverterGUI();
                gui.setVisible(true);
            }
        });
    }
}


