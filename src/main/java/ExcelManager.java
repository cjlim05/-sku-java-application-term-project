import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class ExcelManager {

    private static final LocalDateTime now = LocalDateTime.now();

    // 원하는 형식으로 포맷 지정
    private static final DateTimeFormatter formatter
            = DateTimeFormatter.ofPattern("yyyyMMdd");

    // 포맷 적용하여 문자열 생성
    private static final String fileName = now.format(formatter);



    private static final String FILE_PATH = fileName + ".xlsx";

    // 엑셀 저장 메소드
    public static void saveToExcel(List<String> nameList, List<Integer> quantityList, List<Integer> priceList) {
        try {
            File file = new File(FILE_PATH);
            //엑셀에서 읽고 쓸수 있는 작업, Workbook은 엑셀 파일 전체를 의미한다.
            Workbook workbook;
            Sheet sheet;

            // 엑셀 파일이 존재하면 읽어오고, 없으면 새로 생성
            if (file.exists()) {
                FileInputStream fis = new FileInputStream(file);
                workbook = WorkbookFactory.create(fis); //파일의 내용을 가져와 객체로 만드는 과정
                sheet = workbook.getSheetAt(0);
                fis.close();
            } else {
                workbook = new XSSFWorkbook();//새로운 workbook 객체를 생성
                sheet = workbook.createSheet("Orders"); //엑셀에서 시트 추가
                createHeaderRow(sheet); // 헤더 생성
            }


            // 현재 시간
            String currentTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

            // nameList의 각 항목에 대해 새로운 행을 추가
            for (int i = 0; i < nameList.size(); i++) {
                // 새로운 행 생성 (매 반복마다 새로운 행을 생성)
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);  // 마지막 라인에 칼럼을 추가
                final int totalsum = quantityList.get(i) * priceList.get(i); // 메뉴 총가격 구하기

                // 각 항목을 해당 열에 넣기
                row.createCell(0).setCellValue(currentTime);  // 현재 시간
                row.createCell(1).setCellValue(nameList.get(i));  // 이름
                row.createCell(2).setCellValue(quantityList.get(i));  // 수량
                row.createCell(3).setCellValue(priceList.get(i));  // 가격
                row.createCell(4).setCellValue(totalsum);  // 가격
                System.out.println(nameList.get(i));
            }

            // 엑셀 파일 저장
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            workbook.close();
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 헤더 생성 메소드
    private static void createHeaderRow(Sheet sheet) {
        Row header = sheet.createRow(0);  // 첫 번째 행 생성
        header.createCell(0).setCellValue("Time");    // 시간
        header.createCell(1).setCellValue("Name");    // 이름
        header.createCell(2).setCellValue("Quantity");   // 수량
        header.createCell(3).setCellValue("Price"); // 가격
        header.createCell(4).setCellValue("Total-Price"); // 가격
    }
}
