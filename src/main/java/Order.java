import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.HashMap;
import java.util.Map;

class Order extends JPanel {
    private JPanel ordersPanel; // 여러 주문 항목을 담을 패널
    private JButton reset;
    private JButton purchase;
    private JLabel totalLabel; // 총 가격을 표시할 라벨
    private int totalPrice = 0; // 총 가격을 저장할 변수
    private JLabel timerLabel; // 타이머를 표시할 라벨
    private Timer timer; // 타이머 객체
    private int timeLeft = 100; // 초기 타이머 시간 (100초)

    // 주문 내역 정보를 Map에 저장, 키는 주문 이름, 값은 주문 항목 패널
    private Map<String, OrderItem> orderItems;
    private Ordered getOrder;

    public Order(Ordered getOrder) {
        this.getOrder = getOrder;
        setLayout(new BorderLayout());
        setBackground(Color.lightGray);

        // 주문 항목을 담을 패널 설정
        ordersPanel = new JPanel();
        ordersPanel.setLayout(new BoxLayout(ordersPanel, BoxLayout.Y_AXIS)); // 세로로 쌓이게
        JScrollPane scrollPane = new JScrollPane(ordersPanel);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

        // 주문 내역 타이틀과 주문 패널 추가
        add(new JLabel("주문 내역"), BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);

        // 초기화 버튼 생성 및 이벤트 리스너 추가
        purchase = new JButton("결재하기");
        reset = new JButton("초기화");


        purchase.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                purchase();
            }
        });
        reset.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                resetOrders();
            }
        });

        // 타이머 레이블 설정
        timerLabel = new JLabel("남은 시간: 100초");
        JPanel timerPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        timerPanel.add(timerLabel);
        add(timerPanel, BorderLayout.NORTH);

        // 타이머 초기화
        initializeTimer();

        // 총 가격 표시 패널 설정
        totalLabel = new JLabel("총 가격: 0원");
        JPanel totalPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        totalPanel.add(totalLabel);
        totalPanel.add(purchase);
        totalPanel.add(reset); // 초기화 버튼을 총 가격 패널에 추가
        add(totalPanel, BorderLayout.SOUTH);

        // 주문 항목 저장용 Map 초기화
        orderItems = new HashMap<>();
    }

    // 타이머 초기화 메서드
    private void initializeTimer() {
        timer = new Timer(1000, new ActionListener() { // 1초마다 실행
            @Override
            public void actionPerformed(ActionEvent e) {
                timeLeft--;
                timerLabel.setText("남은 시간: " + timeLeft + "초");
                if (timeLeft <= 0) {
                    timer.stop();
                    resetOrders();
                    JOptionPane.showMessageDialog(null, "시간이 초과되었습니다. 주문이 초기화됩니다.");
                }
            }
        });
        timer.start();
    }

    //타이머 재설정
    private void restartTimer() {
        timeLeft = 100; // 남은 시간을 초기값으로 재설정
        timerLabel.setText("남은 시간: " + timeLeft + "초");
        timer.restart(); // 타이머 다시 시작
    }

    // 주문 내역을 추가하거나 수량을 조절하는 메서드
    public void addOrder(String name, String price) {
        int itemPrice = Integer.parseInt(price.replace("원", "")); // 가격 문자열을 숫자로 변환

        if (orderItems.containsKey(name)) {   //containsKey Map객체에서 특정 키 확인 메서드
            // 이미 존재하는 주문이 들어 왔을때 -> 수량만 변경
            OrderItem orderItem = orderItems.get(name);
            orderItem.incrementQuantity();
            totalPrice += itemPrice;
        } else {
            // 새로운 주문이 들어왔을 때
            OrderItem orderItem = new OrderItem(name, itemPrice);
            orderItems.put(name, orderItem);
            ordersPanel.add(orderItem.getPanel());
            totalPrice += itemPrice;
        }

        updateTotalPrice();
        revalidate(); // 새로 추가된 내용을 갱신
        repaint();
    }

    // 총 가격을 갱신하는 메서드
    private void updateTotalPrice() {
        totalLabel.setText("총 가격: " + totalPrice + "원");
    }

    // 구매 확인 창 메서드
    private void purchase() {
        // 새 JDialog 생성
        JDialog purchaseDialog = new JDialog((Frame) SwingUtilities.getWindowAncestor(this), "주문 확인", true); //팝업창 함수,
        purchaseDialog.setSize(300, 400);
        purchaseDialog.setLayout(new BorderLayout());

        // 최종 주문 확인 창 생성
        JPanel orderSummaryPanel = new JPanel();
        orderSummaryPanel.setLayout(new BoxLayout(orderSummaryPanel, BoxLayout.Y_AXIS));

        // 최종 주문 확인 창에 주문 내역 입력
        for (OrderItem orderItem : orderItems.values()) {
            JPanel itemPanel = new JPanel(new BorderLayout());
            JLabel nameLabel = new JLabel(orderItem.name + " x " + orderItem.quantity);
            JLabel priceLabel = new JLabel(orderItem.price * orderItem.quantity + "원");
            itemPanel.add(nameLabel, BorderLayout.WEST);
            itemPanel.add(priceLabel, BorderLayout.EAST);
            orderSummaryPanel.add(itemPanel);
        }

        // 주문 확인 창에 스크롤 기능 추가
        JScrollPane scrollPane = new JScrollPane(orderSummaryPanel);
        purchaseDialog.add(scrollPane, BorderLayout.CENTER);

        // 총 가격 표시
        JLabel totalLabel = new JLabel("총 가격: " + totalPrice + "원");
        totalLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        purchaseDialog.add(totalLabel, BorderLayout.NORTH);

        // 하단 버튼 패널
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        JButton cancelButton = new JButton("취소");
        JButton confirmButton = new JButton("주문하기");

        // 취소 버튼 이벤트 리스너 -> 구매확인 창만 닫음
        cancelButton.addActionListener(e -> purchaseDialog.dispose());

        // 주문하기 버튼 이벤트 리스너
        confirmButton.addActionListener(e -> {
            getOrder.addMenu(orderItems);
            JOptionPane.showMessageDialog(purchaseDialog, "주문이 완료되었습니다!");
            resetOrders(); // 주문 내역 초기화
            purchaseDialog.dispose();
            restartTimer();
        });

        buttonPanel.add(cancelButton);
        buttonPanel.add(confirmButton);
        purchaseDialog.add(buttonPanel, BorderLayout.SOUTH);

        // 주문 확인 창 표시
        purchaseDialog.setLocationRelativeTo(this); // 창 중앙에 배치
        purchaseDialog.setVisible(true);
    }

    // 주문 내역 초기화 메서드
    private void resetOrders() {
        ordersPanel.removeAll(); // 주문 내역을 초기화
        orderItems.clear(); // 주문 항목 Map도 초기화
        totalPrice = 0; // 총 가격 초기화
        updateTotalPrice(); // 총 가격 라벨 업데이트
        revalidate();
        repaint();
    }





    // 주문 항목 클래스
    class OrderItem {
        private String name;
        private int price;
        private int quantity;
        private JPanel panel; //주문 항목들
        private JLabel quantityLabel;
        private JLabel priceLabel;

        public OrderItem(String name, int price) {
            this.name = name;
            this.price = price;
            this.quantity = 1;

            // 주문 항목 패널 구성
            panel = new JPanel(new BorderLayout()); //패널은 주문들어올때마다 생성되어 인스턴스의 변수이다. 패널을 생설할때마다 동적으로 추가 해줘야함.
            panel.setBackground(Color.white);
            panel.setMinimumSize(new Dimension(300, 100));
            panel.setMaximumSize(new Dimension(300, 100)); //최대 사이즈 조정으로 패널사이즈 조정

            // 주문 내역 레이블 생성
            JLabel nameLabel = new JLabel(name);
            panel.add(nameLabel, BorderLayout.WEST);

            // 수량 조절 버튼 및 수량 레이블 추가
            JPanel controlPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
            JButton minusButton = new JButton("-");
            JButton plusButton = new JButton("+");
            quantityLabel = new JLabel(String.valueOf(quantity));
            priceLabel = new JLabel(price * quantity + "원");

            // 크기 조절
            minusButton.setPreferredSize(new Dimension(40, 30));
            plusButton.setPreferredSize(new Dimension(40, 30));

            minusButton.addActionListener(e -> decrementQuantity());
            plusButton.addActionListener(e -> incrementQuantity());

            controlPanel.add(minusButton);
            controlPanel.add(quantityLabel);
            controlPanel.add(plusButton);

            panel.add(controlPanel, BorderLayout.CENTER);
            panel.add(priceLabel, BorderLayout.EAST);

            // 삭제 버튼 추가
            JButton deleteButton = new JButton("X");
            deleteButton.addActionListener(e -> removeOrderItem());
            panel.add(deleteButton, BorderLayout.SOUTH);
        }
        // Getter 메서드 추가
        public String getName() {
            return name;
        }

        public int getPrice() {
            return price;
        }

        public int getQuantity() {
            return quantity;
        }


        // 수량 증가 메서드
        public void incrementQuantity() {
            quantity++;
            quantityLabel.setText(String.valueOf(quantity));
            priceLabel.setText(price * quantity + "원");
            totalPrice += price;
            updateTotalPrice();
        }

        // 수량 감소 메서드
        public void decrementQuantity() {
            if (quantity > 1) { // 최소 갯수는 1
                quantity--;
                quantityLabel.setText(String.valueOf(quantity));
                priceLabel.setText(price * quantity + "원");
                totalPrice -= price;
                updateTotalPrice();
            }
        }

        // 주문 항목 삭제 메서드
        private void removeOrderItem() {
            ordersPanel.remove(panel);
            totalPrice -= price * quantity;
            orderItems.remove(name);
            updateTotalPrice();
            revalidate();
            repaint();
        }

        // 주문 항목 패널 반환
        // 외부에서도 접근 할 수 있도록 반환 Getter 메서드
        //새로운 주문이 들어 왔을때 새로운 주문을 패널로 추가->이를 동적으로 추가하기 위해 Getter함수 만듦
        public JPanel getPanel() {
            return panel;
        }
    }
}
