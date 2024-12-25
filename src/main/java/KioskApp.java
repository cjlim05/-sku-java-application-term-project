import javax.swing.*;
import java.awt.*;

public class KioskApp {
    private JFrame frame;
    private JPanel mainPanel;
    private JFrame orderedFrame;
    private JPanel orderedPanel;  // 주문 내역 패널
    private Order orderPanel;     // 주문을 처리하는 Order 객체
    private CardLayout cardLayout;

    public KioskApp() {
        // 주문 내역 프레임 설정
        frame = new JFrame("Kiosk");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLayout(new BorderLayout());

        // 주문 받는 프레임 설정
        orderedFrame = new JFrame("Ordered");
        orderedFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        orderedFrame.setSize(800, 600);
        orderedFrame.setLayout(new BorderLayout());

        // 주문 패널 생성
        orderedPanel = new Ordered();         // 주문 내역이 표시될 패널

        JScrollPane orderedPanelScrollPane = new JScrollPane(orderedPanel);
        orderedPanelScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); // 세로 스크롤바 설정
        orderedPanelScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);

        orderedFrame.add(orderedPanelScrollPane);



        //--------------
        orderPanel = new Order((Ordered) orderedPanel); // Order 생성자에 Ordered 패널 전달
        //Order 인스턴스를 생성하고 orderedPanel를 매개변수로 받는데, 이때 orderedPanel의 형이 Ordered임으로 형 변환을 시켜주어야 한다.

        // 주문 내역을 보여주는 프레임에 Ordered 패널 추가
        orderedFrame.add(orderedPanel);


        // Ordered에서 orderedPanel객체를 생성
        // Order에서 orderPanel을 생성, 이때 Ordered클래스의 orderedPanel을 생성자 매개변수로 전달하고 있음

        // 메인 패널과 CardLayout 할당
        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);

        // 커피 및 디저트 패널 생성
        Coffee coffeePanel = new Coffee(orderPanel); // Coffee에서 주문 내역을 처리할 Order 객체 전달
        Dessert dessertPanel = new Dessert(orderPanel);

        // 커피 패널 스크롤 추가
        JScrollPane coffeeScrollPane = new JScrollPane(coffeePanel);
        coffeeScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); // 세로 스크롤바 설정
        coffeeScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED); // 가로 스크롤바 설정

        // 디저트 패널 스크롤 추가
        JScrollPane dessertScrollPane = new JScrollPane(dessertPanel);
        dessertScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); // 세로 스크롤바 설정
        dessertScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED); // 가로 스크롤바 설정

        // 메인 패널에 패널 추가
        mainPanel.add(coffeeScrollPane, "커피");
        mainPanel.add(dessertScrollPane, "디저트");

        // 커피 및 디저트 버튼 생성 및 이벤트 처리
        JButton coffeeButton = new JButton("커피");
        coffeeButton.setFont(new Font("Arial", Font.BOLD, 16)); // 버튼 폰트 스타일 설정
        coffeeButton.setBackground(new Color(75, 192, 192)); // 버튼 배경색 설정
        coffeeButton.setForeground(Color.BLACK); // 버튼 글자색 설정
        coffeeButton.setFocusPainted(false); // 버튼 클릭 시 외곽선 없애기
        coffeeButton.addActionListener(e ->
                cardLayout.show(mainPanel, "커피")
        );

        JButton dessertButton = new JButton("디저트");
        dessertButton.setFont(new Font("Arial", Font.BOLD, 16)); // 버튼 폰트 스타일 설정
        dessertButton.setBackground(new Color(255, 105, 180)); // 버튼 배경색 설정
        dessertButton.setForeground(Color.BLACK); // 버튼 글자색 설정
        dessertButton.setFocusPainted(false); // 버튼 클릭 시 외곽선 없애기
        dessertButton.addActionListener(e ->
                cardLayout.show(mainPanel, "디저트")
        );

        // 버튼 패널 생성
        JPanel buttonPanel = new JPanel();
        buttonPanel.setBackground(new Color(240, 240, 240)); // 배경색 설정
        buttonPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 20, 10)); // 버튼들 간의 간격을 설정
        buttonPanel.add(coffeeButton);
        buttonPanel.add(dessertButton);

        // 메인 프레임에 패널 추가
        frame.add(buttonPanel, BorderLayout.NORTH);
        frame.add(mainPanel, BorderLayout.CENTER);
        frame.add(orderPanel, BorderLayout.EAST);  // 주문 내역 표시

        // 프레임 보이기
        frame.setVisible(true);
        orderedFrame.setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(KioskApp::new); //KioskApp 생성자 실행
        //SwingUtilities.invokeLater() 함수는 EDT에서 스레드를 실행시키는 함수다. UI는 안전하게 업데이트 하기위해서 EDT에서 실행되야 한다
    }
}
