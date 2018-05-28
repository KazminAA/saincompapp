import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.Map;
import java.util.Properties;

public class XlsSainCompApp extends JPanel implements ActionListener {
    private JButton openButton1, openButton2, actionButton, cancelButton;
    private JTextArea log;
    private JFileChooser fc;
    private int sainColumnNumber, nameColumnNumber, numberColumnNumber;

    private File firstFile, secondFile;

    private ElFileParser fileParser = new ElFileParser();

    private XlsSainCompApp() {
        super(new BorderLayout());

        log = new JTextArea(5, 20);
        log.setMargin(new Insets(5, 5, 5, 5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Excell files", "xls", "xlsx"));

        openButton1 = new JButton("Open a First File...",
                createImageIcon("images/Open16.gif"));
        openButton1.addActionListener(this);

        openButton2 = new JButton("Open a Second File...",
                createImageIcon("images/Open16.gif"));
        openButton2.addActionListener(this);

        actionButton = new JButton("Compare");
        actionButton.addActionListener(this);

        cancelButton = new JButton("Cancel");
        cancelButton.addActionListener(this);
        cancelButton.setEnabled(false);

        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton1);
        buttonPanel.add(openButton2);
        buttonPanel.add(actionButton);

        JPanel buttonPanel2 = new JPanel();
        buttonPanel2.add(cancelButton);

        add(buttonPanel, BorderLayout.PAGE_START);
        add(buttonPanel2, BorderLayout.PAGE_END);
        add(logScrollPane, BorderLayout.CENTER);

        setParsedColumns();
    }

    private static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = XlsSainCompApp.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("XLSSAINCOMP");
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new XlsSainCompApp());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            UIManager.put("swing.boldMetal", Boolean.FALSE);
            createAndShowGUI();
        });
    }

    private void fileChoise(JButton button) {
        int returnVal = fc.showOpenDialog(XlsSainCompApp.this);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File selectFile = fc.getSelectedFile();
            if (button.equals(openButton1)) firstFile = selectFile;
            if (button.equals(openButton2)) secondFile = selectFile;
            log.append("Opening: " + selectFile.getName() + ".\n");
            button.setEnabled(false);
        } else {
            log.append("Open command cancelled by user.\n");
        }
        log.setCaretPosition(log.getDocument().getLength());
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == openButton1) {
            fileChoise(openButton1);
            cancelButton.setEnabled(true);
        } else if (e.getSource() == openButton2) {
            fileChoise(openButton2);
            cancelButton.setEnabled(true);
        } else if (e.getSource() == actionButton) {
            fileCompare();
        } else if (e.getSource() == cancelButton) {
            cancel();
            setParsedColumns();
            log.append("All operation cancelled by user.\n");
            log.setCaretPosition(log.getDocument().getLength());
        }

    }

    private void fileCompare() {
        if (firstFile == null || secondFile == null) {
            log.append("Both files must be selected.\n");
            log.setCaretPosition(log.getDocument().getLength());
            return;
        }
        log.append("Comparing... \n");
        log.setCaretPosition(log.getDocument().getLength());
        Map<String, ParsedRow> differanceMap;
        File file = null;
        Workbook result = null;
        try {
            differanceMap = fileParser.getDifference(firstFile, secondFile,
                    new int[]{sainColumnNumber, nameColumnNumber, numberColumnNumber});
            result = new XSSFWorkbook();
            result.createSheet();
            Sheet sheet = result.getSheetAt(0);
            ParsedRow[] rows = new ParsedRow[differanceMap.size()];
            differanceMap.values().toArray(rows);
            for (int i = 0; i < rows.length; i++) {
                sheet.createRow(i);
                Row row = sheet.getRow(i);
                row.createCell(sainColumnNumber);
                row.createCell(nameColumnNumber);
                row.createCell(numberColumnNumber);
                row.getCell(sainColumnNumber).setCellValue(rows[i].getSain());
                row.getCell(nameColumnNumber).setCellValue(rows[i].getName());
                row.getCell(numberColumnNumber).setCellValue(rows[i].getNumber());
            }
        } catch (IOException e) {
            log.append("Something wrong with files!\n");
            log.setCaretPosition(log.getDocument().getLength());
        } catch (InvalidFormatException e) {
            log.append("One from selected file is not an xls/xlsx file.\n");
            log.setCaretPosition(log.getDocument().getLength());
        }
        int returnVal = fc.showSaveDialog(XlsSainCompApp.this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            file = fc.getSelectedFile();
            //This is where a real application would save the file.
            log.append("Saving: " + file.getName() + ".\n");
        } else {
            log.append("Save command cancelled by user.\n");
        }
        log.setCaretPosition(log.getDocument().getLength());
        if (file != null && result != null) {
            try {
                FileOutputStream out = new FileOutputStream(file);
                result.write(out);
                out.close();
            } catch (FileNotFoundException e) {
                log.append("Can't open file to save results.\n");
                log.setCaretPosition(log.getDocument().getLength());
            } catch (IOException e) {
                log.append("Can't save results.\n");
                log.setCaretPosition(log.getDocument().getLength());
            }
        }
        log.append("Done.");
        log.setCaretPosition(log.getDocument().getLength());
        cancel();
    }

    private void cancel() {
        firstFile = null;
        secondFile = null;
        openButton1.setEnabled(true);
        openButton2.setEnabled(true);
        cancelButton.setEnabled(false);
    }

    private void setParsedColumns() {
        Properties properties = new Properties();
        try (FileInputStream inputStream = new FileInputStream("./application.properties");) {
            properties.load(inputStream);
            sainColumnNumber = Integer.valueOf(properties.getProperty("sain.columnnumber", "1"));
            nameColumnNumber = Integer.valueOf(properties.getProperty("name.columnnumber", "2"));
            numberColumnNumber = Integer.valueOf(properties.getProperty("number.columnnumber", "3"));
            log.append("Sain column number set to '" + sainColumnNumber + "'.\n");
            log.append("Name column number set to '" + nameColumnNumber + "'.\n");
            log.append("Quantity column number set to '" + numberColumnNumber + "'.\n");
            log.setCaretPosition(log.getDocument().getLength());
            sainColumnNumber--;
            nameColumnNumber--;
            numberColumnNumber--;
        } catch (Exception e) {
            sainColumnNumber = 0;
            nameColumnNumber = 1;
            numberColumnNumber = 2;
            log.append("Can't load properties. All set to default.\n");
            log.setCaretPosition(log.getDocument().getLength());
        }
    }
}
