import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class XlsSainCompApp extends JPanel implements ActionListener {
    static private final String NEWLINE = "\n";
    private JButton openButton1, openButton2, actionButton;
    private JTextArea log;
    private JFileChooser fc;

    private File firstFile, secondFile;

    public XlsSainCompApp() {
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

        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton1);
        buttonPanel.add(openButton2);
        buttonPanel.add(actionButton);

        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
    }

    protected static ImageIcon createImageIcon(String path) {
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
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new XlsSainCompApp());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                UIManager.put("swing.boldMetal", Boolean.FALSE);
                createAndShowGUI();
            }
        });
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == openButton1) {
            int returnVal = fc.showOpenDialog(XlsSainCompApp.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                firstFile = fc.getSelectedFile();
                log.append("Opening: " + firstFile.getName() + "." + NEWLINE);
            } else {
                log.append("Open command cancelled by user." + NEWLINE);
            }
            log.setCaretPosition(log.getDocument().getLength());

            //Handle save button action.
        } else if (e.getSource() == openButton2) {
            int returnVal = fc.showOpenDialog(XlsSainCompApp.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                secondFile = fc.getSelectedFile();
                log.append("Opening: " + secondFile.getName() + "." + NEWLINE);
            } else {
                log.append("Open command cancelled by user." + NEWLINE);
            }
            log.setCaretPosition(log.getDocument().getLength());
        } else if (e.getSource() == actionButton) {
            log.append("Comparing... \n");
            fileCompare();
            log.setCaretPosition(log.getDocument().getLength());
        }

    }

    private void fileCompare() {
    }
}
