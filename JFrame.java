import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

/**
 * Created by Дмитрий on 22.04.2021.
 */
public class JFrame extends javax.swing.JFrame {

    public JFrame() {
        super("Программа для замены строк в файлах word.");
        this.setBounds(500, 500, 550, 180);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.setLayout(new GridLayout(0, 2));

        JTextField inputCurDir = new JTextField("E:/test.docx", 30);
        JLabel labelCurDir = new JLabel("Имя файла: ");

        JTextField inputStr_input = new JTextField("SORD", 30);
        JLabel labelStr_input = new JLabel("Тег для замены");

        JTextField inputStr_output = new JTextField("VALHALLA", 30);
        JLabel labelStr_output = new JLabel("На что меняем");

        JButton button1 = new JButton("Заменить!");
        JButton button2 = new JButton("Посмотреть результат");

        container.add(labelCurDir);
        container.add(inputCurDir);
        container.add(labelStr_input);
        container.add(inputStr_input);
        container.add(labelStr_output);
        container.add(inputStr_output);


        container.add(button1);
        container.add(button2);


        button1.addActionListener(e -> {

            String filename = inputCurDir.getText();
            String str_input = inputStr_input.getText();
            String str_output = inputStr_output.getText();

            if (checks(inputCurDir.getText()) != 0) {
                setVisible(false);
                JFrame app = new JFrame();
                app.setVisible(true);
                return;
            }
            Replace.replacer(filename, str_input, str_output);

            JOptionPane.showMessageDialog(JFrame.this, "Текст успешно заменен!");

        });

        button2.addActionListener(e -> {

            String filename = inputCurDir.getText();
            String rez = filename.replace(".", "_output.");

            if (checks(rez) != 0) {
                setVisible(false);
                JFrame app = new JFrame();
                app.setVisible(true);
                return;
            }

            String cmd = "c:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.exe " + rez;
            Process p = null;
            try {
                p = Runtime.getRuntime().exec(cmd);
                p.waitFor();
            } catch (IOException e1) {
                e1.printStackTrace();
            } catch (InterruptedException e1) {
                e1.printStackTrace();
            }

        });

    }


    protected int checks(String folder) {

        File f = new File(folder);

        // Проверка на отсутствие введенных данных
        if (folder.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Файл не указан!");
            return 1;
        }

        if (!f.exists()) {
            JOptionPane.showMessageDialog(null, "Указанного файла не существует!");
            return 1;
        }
        return 0;
    }
}