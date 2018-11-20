import javax.swing.*;
import javax.swing.text.html.parser.Parser;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class TestApp extends JFrame{

    JLabel pic;
    Timer tm;
    int x = 0;
    String[] list = { "D:/Workspace/JavaProjects/AdvertTest/src/pictures/1.jpg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/2.jpg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/3.jpg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/4.jpg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/5.jpeg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/6.jpeg",
                      "D:/Workspace/JavaProjects/AdvertTest/src/pictures/7.jpeg",
                    };

        public TestApp() {
            super("Java SlideShow");
            pic = new JLabel();
            Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
            double width = screenSize.getWidth();
            double height = screenSize.getHeight();
            pic.setBounds(0, 0, (int)width, (int)height);
            SetImageSize(6);

            tm = new Timer(2500, new ActionListener(){

                @Override
                public void actionPerformed(ActionEvent e) {
                    SetImageSize(x);
                    x += 1;
                    if (x >= list.length)
                        x = 0;
                }
            });
            add(pic);
            tm.start();
            setLayout(null);
            setSize((int)width, (int)height);
            //DisplayMode.BIT_DEPTH_MULTI;
            //DisplayMode.REFRESH_RATE_UNKNOWN;
            getContentPane().setBackground(Color.decode("#bdb67b"));
            setLocationRelativeTo(null);
            setExtendedState(JFrame.MAXIMIZED_BOTH);
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setUndecorated(true);
            setVisible(true);
        }

        public void SetImageSize(int i){
            ImageIcon icon = new ImageIcon(list[i]);
            Image img = icon.getImage();
            Image newImg = img.getScaledInstance(pic.getWidth(), pic.getHeight(), Image.SCALE_SMOOTH);
            ImageIcon newImc = new ImageIcon(newImg);
            pic.setIcon(newImc);
        }

    public static void main(String[] args){
        new TestApp();
    }
}
