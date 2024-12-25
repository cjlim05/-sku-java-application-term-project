import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.HashMap;
import java.util.Map;

public class Ordered extends JPanel {

    private JPanel mainPanel;  // 주문 내역을 추가할 메인 패널
    private JLabel TotalIncome;
    private int Income = 0;

    List<String> nameList = new ArrayList<>();
    List<Integer> priceList = new ArrayList<>();
    List<Integer> quantityList = new ArrayList<>();


    // 생성자
    public Ordered() {
        // 메인 패널 설정
        setBackground(Color.RED);
        setLayout(new BorderLayout()); // 전체 레이아웃을 BorderLayout으로 설정

        // 총 수입 패널 설정
        TotalIncome = new JLabel("총 매출 : 0원");
        JPanel IncomePanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        IncomePanel.add(TotalIncome);

        // IncomePanel을 맨 위에 추가
        add(IncomePanel, BorderLayout.NORTH);  // 상단에 고정

        // 메인 패널 설정
        mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS)); // 세로로 배치
        JScrollPane scrollPane = new JScrollPane(mainPanel);  // 주문 내역을 스크롤 가능하게 설정
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

        // scrollPane을 중앙에 배치
        add(scrollPane, BorderLayout.CENTER);
    }

    // 주문 내역 추가하는 메서드
    public void addMenu(Map<String, Order.OrderItem> orderItems) {
        // 하나의 주문 세트를 위한 패널 생성
        JPanel orderPanel = new JPanel();
        orderPanel.setLayout(new BorderLayout()); // 주문 내역을 위한 레이아웃 설정
        orderPanel.setBackground(Color.RED);  // 버튼 클릭 전 패널 색상 -> red

        // 총 금액을 누적할 변수 선언
        int totalAmount = 0;


        // StringBuilder - 가변 문자열 처리
        StringBuilder itemText = new StringBuilder();

        // 주문 항목들 추가
        for (Order.OrderItem item : orderItems.values()) {
            itemText.append(item.getName()).append(" : ").append(item.getQuantity()).append("\n");
            totalAmount += item.getPrice() * item.getQuantity(); // 총 금액 계산

            nameList.add(item.getName());
            priceList.add(item.getPrice());
            quantityList.add(item.getQuantity());

            System.out.println(nameList);
            System.out.println(priceList);
            System.out.println(quantityList);

        }
        itemText.append("----------------\n");
        itemText.append("총 금액 : ").append(totalAmount).append("원");


        // JLabel로 줄바꿈 처리
        JLabel itemLabel = new JLabel("<html>" + itemText.toString().replace("\n", "<br>") + "</html>");

        // 주문 내역을 주문 패널에 추가
        orderPanel.add(itemLabel, BorderLayout.CENTER);

        // 총수익 업데이트
        Income += totalAmount;
        updateTotalIncome();

        // 확인 버튼 추가
        JButton actionButton = new JButton("확인");
        actionButton.setPreferredSize(new Dimension(50, 18));

        // 삭제 버튼 추가
        JButton deleteButton = new JButton("삭제");
        deleteButton.setPreferredSize(new Dimension(50, 18));


        actionButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                onActionButtonClick(orderPanel, actionButton, deleteButton);
            }
        });

        deleteButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                onDeleteButtonClick(orderPanel);
            }
        });

        // 버튼 스타일 설정
        actionButton.setForeground(Color.RED);
        actionButton.setOpaque(true);

        deleteButton.setForeground(Color.RED);
        deleteButton.setOpaque(true);

        // 버튼을 주문 패널의 오른쪽에 추가
        JPanel buttonPanel = new JPanel();
        buttonPanel.setLayout(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(actionButton);
        buttonPanel.add(deleteButton);

        orderPanel.add(buttonPanel, BorderLayout.EAST);

        // 주문 패널 테두리 설정
        orderPanel.setBorder(BorderFactory.createLineBorder(Color.BLACK, 2));
        orderPanel.setMaximumSize(new Dimension(1000, 150));
        orderPanel.setAlignmentX(Component.LEFT_ALIGNMENT);

        // 주문 패널을 mainPanel에 맨 위에 추가
        mainPanel.add(orderPanel, 0);  // 0번째 인덱스에 추가하여 가장 위로 배치

        // 주문 패널 갱신
        revalidate(); // 레이아웃 갱신
        repaint(); // 화면 갱신
    }

    private void updateTotalIncome() {

        TotalIncome.setText("총 가격 : " + Income + "원");
    }

    // 버튼 클릭 시 -> 패널 색 변경, 버튼 기능 비활성화
    private void onActionButtonClick(JPanel orderPanel, JButton actionButton, JButton deleteButton) {
        orderPanel.setBackground(Color.GRAY); // 패널 색상 변경
        actionButton.setForeground(Color.GRAY); // 버튼 텍스트 색상 변경
        actionButton.setText("확인됨"); // 버튼 클릭 시 텍스트 변경
        actionButton.setEnabled(false); // 버튼 비활성화
        deleteButton.setEnabled(false); // 버튼 비활성화


        ExcelManager.saveToExcel(nameList, quantityList, priceList);
    }

    // 삭제 버튼 클릭 시 -> 주문 패널 삭제
    private void onDeleteButtonClick(JPanel orderPanel) {
        mainPanel.remove(orderPanel);  // mainPanel에서 해당 주문 패널 삭제
        revalidate(); // 레이아웃 갱신
        repaint(); // 화면 갱신
    }


}
