import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.*;
import java.util.Map;
import java.util.Properties;
import java.util.prefs.Preferences;

public class XlsSainCompApp extends JPanel implements ActionListener {

    private static final int DEFAULT_WIDTH = 450;
    private static final int DEFAULT_HEIGHT = 250;
    private static Preferences root = Preferences.userRoot();
    private static Preferences node = root.node("/com/agrobyd/saincomp");

    private JButton openButton1, openButton2, actionButton, cancelButton;
    private JTextArea log;
    private JFileChooser fc;
    private int operationColumnNumber, sainColumnNumber, nameColumnNumber, numberColumnNumber, notesColumnNumber, skipFirstRows;

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


    private static void createAndShowGUI() {
        JFrame frame = new JFrame("XLSSAINCOMP");
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new XlsSainCompApp());

        frame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                node.putInt("left", frame.getX());
                node.putInt("top", frame.getY());
                super.windowClosing(e);
            }
        });

        int left = node.getInt("left", 0);
        int top = node.getInt("top", 0);

        //Display the window.
        frame.setBounds(left, top, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        frame.setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            UIManager.put("swing.boldMetal", Boolean.FALSE);
            createAndShowGUI();
        });
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
                    new int[]{sainColumnNumber, nameColumnNumber, numberColumnNumber, notesColumnNumber, skipFirstRows, operationColumnNumber});
            result = new XSSFWorkbook();
            result.createSheet();
            CellStyle style1 = result.createCellStyle();
            Font font1 = result.createFont();
            font1.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            style1.setFont(font1);
            CellStyle style2 = result.createCellStyle();
            Font font2 = result.createFont();
            font2.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
            style2.setFont(font2);
            Sheet sheet = result.getSheetAt(0);
            for (int i = 0; i <= skipFirstRows; i++) {
                sheet.createRow(i);
            }
            Row row = sheet.getRow(5);
            row.createCell(operationColumnNumber);
            row.createCell(sainColumnNumber);
            row.createCell(nameColumnNumber);
            row.createCell(numberColumnNumber);
            row.createCell(notesColumnNumber);
            row.getCell(operationColumnNumber).setCellValue("Операция");
            row.getCell(sainColumnNumber).setCellValue("Спецификация");
            row.getCell(nameColumnNumber).setCellValue("Материал");
            row.getCell(numberColumnNumber).setCellValue("Кол-во");
            row.getCell(notesColumnNumber).setCellValue("Примечания технолога");
            ParsedRow[] rows = new ParsedRow[differanceMap.size()];
            differanceMap.values().toArray(rows);
            for (int i = 6; i < rows.length + 6; i++) {
                sheet.createRow(i);
                row = sheet.getRow(i);
                row.createCell(operationColumnNumber);
                row.createCell(sainColumnNumber);
                row.createCell(nameColumnNumber);
                row.createCell(numberColumnNumber);
                row.createCell(notesColumnNumber);
                row.getCell(operationColumnNumber).setCellValue(rows[i - 6].getOperation());
                row.getCell(sainColumnNumber).setCellValue(rows[i - 6].getSain());
                row.getCell(nameColumnNumber).setCellValue(rows[i - 6].getName());
                row.getCell(numberColumnNumber).setCellValue(rows[i - 6].getNumber());
                if (rows[i - 6].getNumber() < 0) {
                    row.getCell(numberColumnNumber).setCellStyle(style1);
                } else {
                    row.getCell(numberColumnNumber).setCellStyle(style2);
                }
                row.getCell(notesColumnNumber).setCellValue(rows[i - 6].getNotes());
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
            operationColumnNumber = Integer.valueOf(properties.getProperty("operation.columnnumber", "0"));
            sainColumnNumber = Integer.valueOf(properties.getProperty("sain.columnnumber", "1"));
            nameColumnNumber = Integer.valueOf(properties.getProperty("name.columnnumber", "2"));
            numberColumnNumber = Integer.valueOf(properties.getProperty("number.columnnumber", "3"));
            notesColumnNumber = Integer.valueOf(properties.getProperty("notes.columnnumber", "4"));
            skipFirstRows = Integer.valueOf(properties.getProperty("skip.first.rows", "4"));
            log.append("Operation column number set to '" + operationColumnNumber + "'.\n");
            log.append("Sain column number set to '" + sainColumnNumber + "'.\n");
            log.append("Name column number set to '" + nameColumnNumber + "'.\n");
            log.append("Quantity column number set to '" + numberColumnNumber + "'.\n");
            log.append("Notes column number set to '" + notesColumnNumber + "'.\n");
            log.append("Skip First Rows number set to '" + skipFirstRows + "'.\n");
            log.setCaretPosition(log.getDocument().getLength());
            operationColumnNumber--;
            sainColumnNumber--;
            nameColumnNumber--;
            numberColumnNumber--;
            notesColumnNumber--;
        } catch (Exception e) {
            operationColumnNumber = 0;
            sainColumnNumber = 1;
            nameColumnNumber = 2;
            numberColumnNumber = 3;
            notesColumnNumber = 4;
            skipFirstRows = 5;
            log.append("Can't load properties. All set to default.\n");
            log.setCaretPosition(log.getDocument().getLength());
        }
    }
}
