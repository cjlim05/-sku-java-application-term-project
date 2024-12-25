import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

class Dessert extends JPanel {
    private Order orderedPanel;

    // Ordered 객체를 전달받는 생성자
    public Dessert(Order orderedPanel) {
        this.orderedPanel = orderedPanel;
        setLayout(new FlowLayout(FlowLayout.LEFT, 10, 10));
        setBackground(Color.PINK);

        // 디저트 메뉴 항목 추가
        addMenuItem("베이글", "5000원", "bagel.png");
        addMenuItem("마카롱", "3000원", "macaron.png");
        addMenuItem("브라우니", "3500원", "brownie.png");
        addMenuItem("쿠키", "2000원", "cookie.png");

        setPreferredSize(new Dimension(500, 1500));
    }

    // 메뉴 항목을 추가하는 메서드
    private void addMenuItem(String name, String price, String imagePath) {
        JPanel itemPanel = new JPanel();
        itemPanel.setLayout(new BorderLayout());
        itemPanel.setPreferredSize(new Dimension(150, 200));
        itemPanel.setBackground(new Color(240, 240, 240)); // 밝은 회색 배경
        itemPanel.setBorder(BorderFactory.createLineBorder(Color.GRAY, 1)); // 테두리 추가

        // 이미지 로드 및 크기 조정
        ImageIcon originalIcon = new ImageIcon(getClass().getResource("/images/" + imagePath));
        Image image = originalIcon.getImage();
        Image resizedImage = image.getScaledInstance(150, 150, Image.SCALE_SMOOTH);
        ImageIcon resizedIcon = new ImageIcon(resizedImage);

        JLabel imageLabel = new JLabel();
        imageLabel.setIcon(resizedIcon);
        imageLabel.setHorizontalAlignment(SwingConstants.CENTER);

        // 이름 및 가격 라벨 꾸미기
        JLabel nameLabel = new JLabel(name, SwingConstants.CENTER);
        nameLabel.setFont(new Font("Arial", Font.BOLD, 14)); // 굵은 글씨
        nameLabel.setForeground(new Color(50, 50, 50)); // 진한 회색

        JLabel priceLabel = new JLabel(price, SwingConstants.CENTER);
        priceLabel.setFont(new Font("Arial", Font.PLAIN, 12)); // 일반 글씨
        priceLabel.setForeground(new Color(100, 100, 100)); // 연한 회색

        // 마우스 호버 효과 추가
        itemPanel.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                itemPanel.setBackground(new Color(220, 220, 220)); // 호버 시 밝은 회색
                itemPanel.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)); // 손 모양 커서
            }

            @Override
            public void mouseExited(MouseEvent e) {
                itemPanel.setBackground(new Color(240, 240, 240)); // 기본 색상
            }
        });

        itemPanel.add(imageLabel, BorderLayout.CENTER);
        itemPanel.add(nameLabel, BorderLayout.NORTH);
        itemPanel.add(priceLabel, BorderLayout.SOUTH);

        // 클릭 리스너
        itemPanel.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                JOptionPane.showMessageDialog(null, name + "를(을) 선택하셨습니다.");
                processOrder(name, price);
            }
        });

        add(itemPanel);
    }


    // 주문 처리 메서드
    public void processOrder(String name, String price) {
        if (orderedPanel != null) {
            orderedPanel.addOrder(name, price); // 주문을 처리하는 메서드
        } else {
            System.err.println("Ordered panel is null.");
        }
    }
}
