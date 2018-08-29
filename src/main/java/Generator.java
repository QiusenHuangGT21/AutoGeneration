import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleFloatProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import jdk.jfr.StackTrace;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Array;
import java.math.BigDecimal;
import java.nio.file.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Predicate;

import javafx.application.Application;

public class Generator extends Application {

    private static File workbookFile;
    private static Map<String, ArrayList<Car>> dataMap = new HashMap();
    private static Map<String, ArrayList<Car>> allData = new HashMap();
    private static ObservableList<String> catalogTime = FXCollections.observableArrayList();
    private static String date;
    private static String dateFileName;
    private static boolean isDateFileNameRegistered = false;
    private static Workbook workbook;
    private static SimpleIntegerProperty stageWidth = new SimpleIntegerProperty(0);
    private static final String[] headers
            = {"序号", "日期", "产品名称", "重量（吨）", "单价(元/吨)", "金额（元）"
            , "车辆号码","备注"};
    private static String monthInFile;
    private static DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
    private static SimpleStringProperty selectedCustomer = new SimpleStringProperty("未选择客户");
    private static SimpleFloatProperty tableTotal = new SimpleFloatProperty(0);
    private static SimpleFloatProperty tableWeightTotal = new SimpleFloatProperty(0);
    private static SimpleBooleanProperty progressVisibility = new SimpleBooleanProperty(false);
    private static ArrayList<String> m31 = new ArrayList<>(Arrays.asList("01", "03", "05", "07", "08", "10", "12"));
    private static Map<String, ArrayList<Total>> customerTotal = new HashMap<>();
    private static Total selectedTotal = new Total();
    private static boolean isWindows = System.getProperty("os.name").toLowerCase().contains("windows");

    private static void saveFiles() {
        File file = new File("./data/" + dateFileName + ".txt");
        if (file.exists()) {
            file.delete();
        }
        Path path = Paths.get("./data/" + dateFileName + ".txt");
        try {
            Files.write(path, (String.format("+%s", dateFileName + "\n")).getBytes());
        } catch (IOException e) {
            errorReport("创建" + dateFileName + "时发生错误");
        }

        for (String c : dataMap.keySet()) {
            try{
                Files.write(path, ("^" + c + "\n").getBytes() , StandardOpenOption.APPEND);
            } catch (IOException e) {
                errorReport("保存" + c + "时发生错误");
            }
            for (Car car : dataMap.get(c)) {
                try {
                    Files.write(path, car.toString().getBytes(), StandardOpenOption.APPEND);
                } catch (IOException e) {
                    errorReport("保存" + c + "时发生错误");
                }
            }
        }
    }

    private static void exitSave() {
        File file;
        if (isWindows) {
            file = new File(".\\data\\");
        } else {
            file = new File("./data/");
        }
        File[] files = file.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.endsWith(".txt");
            }
        });
        for (File f : files) {
            f.delete();
        }
        Map<String, Map<String, ArrayList<Car>>> timeBasedMap = new HashMap<>();
        for (String customer : allData.keySet()) {
            for (Car car : allData.get(customer)) {
                if (!timeBasedMap.containsKey(car.getDate().substring(0,7))) {
                    timeBasedMap.put(car.getDate().substring(0,7), new HashMap<String, ArrayList<Car>>());
                }
                if (!timeBasedMap.get(car.getDate().substring(0,7)).containsKey(customer)) {
                    timeBasedMap.get(car.getDate().substring(0,7)).put(customer, new ArrayList<Car>());
                }
                timeBasedMap.get(car.getDate().substring(0,7)).get(customer).add(car);
            }
        }

        for (String time : timeBasedMap.keySet()) {
            Path p;
            if (!isWindows) {
                p = Paths.get("./data/" + time.split("/")[0]
                        + "年 " + time.split("/")[1] + " 月.txt");
            } else {
                p = Paths.get(".\\data\\" + time.split("/")[0]
                        + "年 " + time.split("/")[1] + " 月.txt");
            }
            try {
                Files.write(p
                        , (String.format("+" + time.split("/")[0] + "年 " + time.split("/")[1] + " 月\n")
                                .getBytes()));
            } catch (IOException e) {
                errorReport("保存文件时发生错误");
            }
            for (String customer : timeBasedMap.get(time).keySet()) {
                try{
                    Files.write(p, ("^" + customer + "\n").getBytes() , StandardOpenOption.APPEND);
                } catch (IOException e) {
                    errorReport("保存" + customer + "时发生错误");
                }
                for (Car car : timeBasedMap.get(time).get(customer)) {
                    try {
                    Files.write(p, car.toString().getBytes(), StandardOpenOption.APPEND);
                } catch (IOException e) {
                    errorReport("保存\n" + customer + "\n" + car + "时发生错误");
                }
                }
            }
        }
    }

    private static void readFiles() {
        File data;
        if (!isWindows) {
            data = new File("./data");
        } else {
            data = new File(".\\data");
        }
        File[] dataMonth = data.listFiles();
        String customer = "";
        for (File c : dataMonth) {
            try {
                Scanner scanner = new Scanner(c);
                scanner.useDelimiter("\n");
                while (scanner.hasNext()) {
                    String line = scanner.next();
                    if (line.startsWith("+")) {
                        //start of file of the month
                    } else if (line.startsWith("^")){
                        //start of one customer
                        customer = line.substring(1);
                        if (!allData.containsKey(customer)) {
                            allData.put(customer, new ArrayList<Car>());
                        }
                    } else if (line.startsWith("2")) {
                        //start of an entry
                        if (!catalogTime.contains(line.substring(0, 7))) {
                            catalogTime.add(line.substring(0, 7));
                        }
                        Car carR = new Car(line);
                        allData.get(customer).add(carR);
                    }
                }
            } catch (Exception e) {
                errorReport("读取" + c.getName() + "时出现错误");
            }
        }
        for (String c : allData.keySet()) {
            customerTotal.put(c, new ArrayList<>());
            ArrayList<Total> totalArrayList = customerTotal.get(c);
            for (String time : catalogTime) {
                totalArrayList.add(new Total(time));
            }
            for (Car car : allData.get(c)) {
                for (Total t : customerTotal.get(c)) {
                    if (car.getDate().substring(0, 7).equals(t.getDate())) {
                        t.increment(car);
                    }
                }
            }
            for (Total t : customerTotal.get(c)) {
                t.setRmbTotal();
            }
        }
    }


    private static void scan() {
        dataMap.clear();
        try {
            workbook = WorkbookFactory.create(workbookFile);
        } catch (IOException e) {
            errorReport("IOException");
        } catch (InvalidFormatException e) {
            errorReport("InvalidFormat");
        }

        int sheetNum = workbook.getNumberOfSheets();

        for (int index = sheetNum - 1; index >= 0; index--) {
            try {
                boolean start = false;
                Sheet sheet = workbook.getSheetAt(index);
                DataFormatter dataFormatter = new DataFormatter();
                int rowNum = 0;
                for (Row row : sheet) {
                    rowNum++;
                    if ((rowNum == 2) && !isDateFileNameRegistered) {
                        dateFileName = dataFormatter.formatCellValue(row.getCell(0));
                        isDateFileNameRegistered = true;
                    }
                    String carNum = "";
                    String brandName = "";
                    Float weight = new Float(0);
                    for (int i = 0; i <= 6; i++) {
                        String cellValue = dataFormatter.formatCellValue(row.getCell(i));
                        cellValue = cellValue.replace(" ", "_");
                        if (cellValue.equals("1") && !start
                                && !dataFormatter.formatCellValue(row.getCell(i + 1))
                                .equals("END")) {
                            start = true;
                        } else if (cellValue.equals("END")) {
                            start = false;
                        }
                        if (start) {
                            switch (i) {
                                case 1:
                                    if (cellValue.equals("")) {
                                    } else {
                                        Date d = row.getCell(i).getDateCellValue();
                                        date = dateFormat.format(d);
                                    }
                                    break;
                                case 2:
                                    carNum = cellValue;
                                    break;
                                case 3:
                                    brandName = cellValue;
                                    break;
                                case 4:
                                    weight = new Float(cellValue);
                                    break;
                                case 5:
                                    if (dataMap.containsKey(cellValue)) {
                                    } else {
                                        dataMap.put(cellValue, new ArrayList<Car>());
                                    }
                                    dataMap.get(cellValue).add(
                                            new Car(date, brandName, weight, carNum, cellValue)
                                    );
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            } catch (Exception e) {
                errorReport(String.format("读取%d号工作表时发生错误", index));
            }

        }
    }

    private static void outputSheets(String timeString) throws IOException{
//        for (String s : allData.keySet()) {
//            System.out.println(s);
//            for (Car c ; allData.get(s)) {
//                System.out.println("\t" + c);
//            }
//        }
        progressVisibility.set(true);
        String[] keys = allData.keySet().toArray(new String[allData.keySet().size()]);
        dateFileName = timeString.substring(0, 4) + "年 " + timeString.substring(5, 7)
                + " 月";
        if (!isWindows) {
            new File("./输出/" + dateFileName + "结算单").mkdir();
        } else {
            new File(".\\输出\\" + dateFileName + "结算单").mkdir();
        }
        for (int j = 0; j < keys.length; j++) {
            for (Car car : allData.get(keys[j])) {
                if (car.getDate().substring(0, 7).equals(timeString)) {
                    Workbook output = new XSSFWorkbook();
                    Sheet outputSheet = output.createSheet(keys[j]);

                    Font titleFont = output.createFont();
                    titleFont.setFontHeightInPoints((short) 20);
                    CellStyle titleCellStyle = output.createCellStyle();
                    titleCellStyle.setFont(titleFont);

                    Cell cell1 = outputSheet.createRow(0).createCell(0);
                    cell1.setCellValue("结算单");
                    cell1.setCellStyle(titleCellStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(0,0,0,7));
                    CellUtil.setAlignment(cell1, output, CellStyle.ALIGN_CENTER);
                    Cell cell2 = outputSheet.createRow(1).createCell(0);
                    cell2.setCellValue(dateFileName);
                    Font timeFont = output.createFont();
                    timeFont.setFontHeightInPoints((short) 12);
                    CellStyle timeStyle = output.createCellStyle();
                    timeStyle.setFont(timeFont);
                    cell2.setCellStyle(timeStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(1,1,0,7));
                    CellUtil.setAlignment(cell2, output, CellStyle.ALIGN_CENTER);

                    CellStyle contentStyle = output.createCellStyle();
                    Font contentFont = output.createFont();
                    contentFont.setFontHeightInPoints((short) 12);
                    contentStyle.setFont(contentFont);
                    outputSheet.createRow(2);
                    Cell cell3 = outputSheet.createRow(2).createCell(0);
                    cell3.setCellValue("客户：" + keys[j]);
                    outputSheet.addMergedRegion(new CellRangeAddress(2,2,0,5));
                    CellUtil.setAlignment(cell3, output, CellStyle.ALIGN_LEFT);

                    CellStyle dataStyle = contentStyle;
                    dataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    dataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    dataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
                    dataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                    Row r3 = outputSheet.createRow(3);
                    Cell cell4 = r3.createCell(0);
                    cell4.setCellValue("您本月提货付款情况如下:");
                    outputSheet.addMergedRegion(new CellRangeAddress(3,3,0,3));
                    CellUtil.setAlignment(cell4, output, CellStyle.ALIGN_LEFT);
                    Cell cell5 = r3.createCell(4);

                    outputSheet.setColumnWidth(0,1200);
                    outputSheet.setColumnWidth(1, 3000);
                    outputSheet.setColumnWidth(2, 3000);
                    outputSheet.setColumnWidth(3, 2500);
                    outputSheet.setColumnWidth(4, 2800);
                    outputSheet.setColumnWidth(5, 3000);
                    outputSheet.setColumnWidth(6, 2500);

                    outputSheet.createRow(4).createCell(0).setCellValue("提货明细:");

                    Row r5 = outputSheet.createRow(5);
                    int dataRowStart = 6;

                    for (int i = 0; i < headers.length; i++) {
                        Cell headerCell = r5.createCell(i);
                        headerCell.setCellValue(headers[i]);
                        headerCell.setCellStyle(dataStyle);
                        CellUtil.setAlignment(headerCell, output, CellStyle.ALIGN_CENTER);
                    }
                    int num = 1;
                    int totalWeight = 0;
                    for (Car c : allData.get(keys[j])) {
                        if (c.getDate().substring(0, 7).equals(timeString)) {
                            Row newRow = outputSheet.createRow(dataRowStart++);
                            Cell numberCell = newRow.createCell(0);
                            numberCell.setCellValue(num++);
                            numberCell.setCellStyle(dataStyle);
                            CellUtil.setAlignment(numberCell,output, CellStyle.ALIGN_CENTER);
                            Cell dateCell = newRow.createCell(1);
                                dateCell.setCellValue(c.getDate());
                                dateCell.setCellStyle(dataStyle);
                            Cell brandCell = newRow.createCell(2);
                            brandCell.setCellValue(c.getBrandName());
                            brandCell.setCellStyle(dataStyle);
                            Cell weightCell = newRow.createCell(3);
                            weightCell.setCellValue(c.getWeight());
                            weightCell.setCellStyle(dataStyle);
                            CellUtil.setAlignment(weightCell, output, CellStyle.ALIGN_CENTER);
                            newRow.createCell(4).setCellStyle(dataStyle);
                            Cell carPrice = newRow.createCell(5);
                            carPrice.setCellFormula(String.format("D%d*E%d", num + 5, num + 5));
                            carPrice.setCellStyle(dataStyle);
                            Cell carNumberCell = newRow.createCell(6);
                            Cell blankCell = newRow.createCell(7);
                            blankCell.setCellStyle(dataStyle);
                            carNumberCell.setCellValue((" " + c.getCarNumber()));
                            carNumberCell.setCellStyle(dataStyle);
                            CellUtil.setAlignment(carNumberCell, output, CellStyle.ALIGN_CENTER);
                            totalWeight += c.getWeight();
                        }
                    }
                    Row total = outputSheet.createRow(num + 5);
                    Cell hjCell = total.createCell(0);
                    hjCell.setCellValue("合计");
                    hjCell.setCellStyle(dataStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(num + 5, num + 5, 0, 2));
                    CellUtil.setAlignment(hjCell, output, CellStyle.ALIGN_CENTER);
                    Cell blankCell = total.createCell(1);
                    blankCell.setCellStyle(dataStyle);
                    Cell blankCell1 = total.createCell(2);
                    blankCell1.setCellStyle(dataStyle);
                    Cell totalCell = total.createCell(3);
                    totalCell.setCellFormula(String.format("SUM(D7:D%d)", num + 5));
                    totalCell.setCellStyle(dataStyle);
                    CellUtil.setAlignment(totalCell, output, CellStyle.ALIGN_CENTER);
                    Cell blankCell2 = total.createCell(4);
                    blankCell2.setCellStyle(dataStyle);
                    Cell blankCell3 = total.createCell(5);
                    blankCell3.setCellFormula(String.format("SUM(F7:F%d)", num + 5));
                    blankCell3.setCellStyle(dataStyle);
                    Cell blankCell4 = total.createCell(6);
                    blankCell4.setCellStyle(dataStyle);
                    Cell blankCell5 = total.createCell(7);
                    blankCell5.setCellStyle(dataStyle);
                    String lastDay = "";
                    if (m31.contains(timeString.substring(5, 7))) {
                        lastDay = "31";
                    } else if (timeString.substring(5, 7).equals("02")) {
                       int year = Integer.parseInt(timeString.substring(0, 4));
                        if (year % 4 == 0) {
                            lastDay = "29";
                        } else {
                            lastDay = "28";
                        }
                    } else {
                        lastDay = "30";
                    }
                    cell5.setCellValue(String.format("结算时间：%s1日至%s%s日"
                            , dateFileName.replace(" ", "")
                            , dateFileName.replace(" ", ""), lastDay));
                    Row totalRow = outputSheet.createRow(num + 6);
                    Cell rmbChineseCell = totalRow.createCell(0);
                    rmbChineseCell.setCellValue("本期货款： 大写人民币 ");
                    rmbChineseCell.setCellStyle(dataStyle);
                    for (int i = 1; i < 8; i++) {
                        totalRow.createCell(i).setCellStyle(dataStyle);
                    }
                    outputSheet.addMergedRegion(new CellRangeAddress(rmbChineseCell.getRow().getRowNum()
                            , rmbChineseCell.getRow().getRowNum(), 0,7));

                    Row fkqkRow = outputSheet.createRow(num + 7);
                    fkqkRow.createCell(0).setCellValue("付款情况");

                    Row lastMonthDebtRow = outputSheet.createRow(num + 8);
                    lastMonthDebtRow.createCell(0).setCellValue("上月末欠款");
                    lastMonthDebtRow.getCell(0).setCellStyle(dataStyle);
                    lastMonthDebtRow.createCell(1).setCellStyle(dataStyle);
                    lastMonthDebtRow.createCell(2).setCellStyle(dataStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(lastMonthDebtRow.getRowNum()
                            , lastMonthDebtRow.getRowNum(), 0, 1));

                    Row thisMonthMoney = outputSheet.createRow(num + 9);
                    thisMonthMoney.createCell(0).setCellValue("本月货款");
                    thisMonthMoney.getCell(0).setCellStyle(dataStyle);
                    thisMonthMoney.createCell(1).setCellStyle(dataStyle);
                    thisMonthMoney.createCell(2).setCellStyle(dataStyle);
                    thisMonthMoney.getCell(2).setCellFormula(String.format("SUM(F7:F%d)", num + 5));
                    outputSheet.addMergedRegion(new CellRangeAddress(thisMonthMoney.getRowNum()
                            , thisMonthMoney.getRowNum(), 0, 1));

                    Row thisMonthPaid = outputSheet.createRow(num + 10);
                    thisMonthPaid.createCell(0).setCellValue("本月收到货款");
                    thisMonthPaid.getCell(0).setCellStyle(dataStyle);
                    thisMonthPaid.createCell(1).setCellStyle(dataStyle);
                    thisMonthPaid.createCell(2).setCellStyle(dataStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(thisMonthPaid.getRowNum()
                            , thisMonthPaid.getRowNum(), 0, 1));

                    Row thisMonthDebt = outputSheet.createRow(num + 11);
                    thisMonthDebt.createCell(0).setCellValue("本月末欠货款");
                    thisMonthDebt.getCell(0).setCellStyle(dataStyle);
                    thisMonthDebt.createCell(1).setCellStyle(dataStyle);
                    thisMonthDebt.createCell(2).setCellStyle(dataStyle);
                    outputSheet.addMergedRegion(new CellRangeAddress(thisMonthDebt.getRowNum()
                            , thisMonthDebt.getRowNum(), 0, 1));

                    outputSheet.createRow(num + 12).createCell(0)
                            .setCellValue("  以上货款双方核对无误， 请及时按月付清货款");

                    Row rowA = outputSheet.createRow(num + 13);
                    rowA.createCell(0).setCellValue("收款户名： 福建邦信贸易有限公司");
                    rowA.createCell(4).setCellValue("收款户名： 黄海梅");

                    Row rowB = outputSheet.createRow(num + 14);
                    rowB.createCell(0).setCellValue("开户行： 建设银行龙岩南城支行");
                    rowB.createCell(4).setCellValue("开户行： 建设银行龙岩南城支行");

                    Row rowC = outputSheet.createRow(num + 15);
                    rowC.createCell(0).setCellValue("账号： 3505 0169 7753 0000 0060");
                    rowC.createCell(4).setCellValue("卡号： 6217 0018 8000 6473 076");

                    Row rowD = outputSheet.createRow(num + 16);
                    rowD.createCell(0).setCellValue("供货方");
                    rowD.createCell(5).setCellValue("购货方");

                    Row rowE = outputSheet.createRow(num + 17);
                    rowE.createCell(0).setCellValue("经办人：");
                    rowE.createCell(5).setCellValue("经办人：");

                    Row rowF = outputSheet.createRow(num + 18);
                    rowF.createCell(1).setCellValue("年       月       日");
                    rowF.createCell(6).setCellValue("年       月       日");

                    FileOutputStream fileOut;
                    if (!isWindows) {
                        fileOut = new FileOutputStream("./输出/" + dateFileName
                                + "结算单/" + keys[j] + dateFileName + "结算单" + ".xlsx");
                    } else {
                        fileOut = new FileOutputStream(".\\输出\\" + dateFileName
                                + "结算单\\" + keys[j] + dateFileName + "结算单" + ".xlsx");
                    }

                    output.write(fileOut);
                    fileOut.close();
                }
            }
        }
    }


    private static void errorReport(String msg) {
        Stage stage = new Stage();
        Pane pane = new VBox();
        Scene scene = new Scene(pane);
        Label text = new Label("未知错误");
        Label label = new Label(msg);
        Button confirm = new Button("确认");
        confirm.setOnAction(event -> {
            stage.close();
        });
        Button exit = new Button("退出");
        exit.setOnAction(event -> System.exit(0));
        pane.getChildren().addAll(text, label, exit);
        stage.setScene(scene);
        stage.show();
    }

    private static void openDialog(String msg, EventHandler eventHandler) {
        Stage stage = new Stage();
        Pane pane = new VBox();
        Scene scene = new Scene(pane);
        Label label = new Label(msg);
        Button exit = new Button("取消");
        Button confirm = new Button("确定");
        confirm.setOnMousePressed(eventHandler);
        confirm.setOnMouseReleased(event -> stage.close());
        exit.setOnAction(event -> stage.close());
        pane.getChildren().addAll(label, confirm, exit);
        stage.setScene(scene);
        stage.show();
    }

    private static void outputDialog() {
        catalogTime.sort(new Comparator<String>() {
            @Override
            public int compare(String o1, String o2) {
                if (Integer.parseInt(o1.substring(0, 4)) == Integer.parseInt(o2.substring(0, 4))) {
                    return Integer.parseInt(o1.substring(5, 7)) - Integer.parseInt(o2.substring(5, 7));
                }
                return Integer.parseInt(o1.substring(0, 4)) - Integer.parseInt(o2.substring(0, 4));
            }
        });

        Stage stage = new Stage();
        Pane pane = new VBox();
        Scene scene = new Scene(pane);
        ListView<String> months = new ListView<>();
        SimpleStringProperty selectedTime = new SimpleStringProperty("未选择");
        months.setItems(catalogTime);
        months.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                selectedTime.set(newValue);
            }
        });
        Label label = new Label();
        label.textProperty().bind(selectedTime);
        Pane hb = new HBox();
        Button output = new Button("输出");
        Button cancel = new Button("取消");
        output.setOnAction(event -> {
            System.out.println(selectedTime.get());
            if (selectedTime.getName().equals("未选择")) {
                errorReport("请选择月份");
            } else {
                try {
                    outputSheets(selectedTime.get());
                } catch (IOException e) {
                    errorReport("生成有误");
                    e.printStackTrace();
                }
            };
        });
        cancel.setOnAction(event -> {
            stage.close();
        });
        Label progress = new Label();
        Label perLabel = new Label("%");
        Pane perPane = new HBox();
        perPane.getChildren().addAll(progress, perLabel);
        hb.getChildren().addAll(output, cancel);
        perPane.visibleProperty().bind(progressVisibility);
        pane.getChildren().addAll(months, label, perPane, hb);
        stage.setScene(scene);
        stage.show();
        stage.setOnCloseRequest(event -> progressVisibility.set(false));
    }

    private static void alldataChecker(String msg) {
        System.out.println(allData.get("永杭5标（赵立宇）") + msg);
    }

    @Override
    public void start(Stage primaryStage) throws InvalidFormatException, OpenXML4JRuntimeException {
        if (!isWindows) {
            new File("./输出").mkdir();
            new File("./data").mkdir();
        } else {
            new File(".\\输出").mkdir();
            new File(".\\data").mkdir();
        }
        primaryStage.setTitle("结算单管理工具");
        primaryStage.setMinHeight(950);
        primaryStage.setMinWidth(1500);
        primaryStage.setOnCloseRequest(event -> {
            exitSave();
        });
        stageWidth.bind(primaryStage.widthProperty());
        Pane mainPane = new VBox();
        Scene mainScene = new Scene(mainPane);
        primaryStage.setScene(mainScene);
        primaryStage.show();
//        primaryStage.setFullScreen(true);
        //1st Pane
        MenuBar menuBar = new MenuBar();
        //Menu, 2nd
        Pane fileBox = new HBox();
        //for file process info, 2nd
        Pane dataViewBox = new HBox();
        ListView<String> customersListView = new ListView<>();
        customersListView.setPrefSize(400, 900);
        TableView<Car> selectedTableView = new TableView<>();
        Pane tablePane = new VBox();
        //for table summery
        Pane summeryPane = new GridPane();
        selectedTableView.setPrefSize(800, 800);
        tablePane.getChildren().addAll(selectedTableView, summeryPane);
        dataViewBox.getChildren().addAll(customersListView, tablePane);
        //for data
        Pane statuesBar = new StackPane();
        //for statues
        mainPane.getChildren().addAll(menuBar, dataViewBox, statuesBar);
        Label monthlySummeryCNY = new Label("月合计");
        Pane monthSummeryPane = new VBox();
        ComboBox<String> monthComboBox = new ComboBox<>();
        Pane monthlySummeryGridPane = new GridPane();
        Label totalWeightMonthlyCHN = new Label("重量总计（吨）：");
        GridPane.setConstraints(totalWeightMonthlyCHN,0,1);
        Label totalWeightMonthlyLabel = new Label();
        GridPane.setConstraints(totalWeightMonthlyLabel,1,1);
        Label totalMonthlyCHN = new Label("金额合计（元）：");
        GridPane.setConstraints(totalMonthlyCHN,0,2);
        Label tableTotalMonthlyLabel = new Label();
        GridPane.setConstraints(tableTotalMonthlyLabel,1,2);
        Label totalInCNYMonthlyCHN = new Label("大写人民币：");
        GridPane.setConstraints(totalInCNYMonthlyCHN, 0,3);
        Label totalInCNYMonthlyLabel = new Label();
        GridPane.setConstraints(totalInCNYMonthlyLabel,1,3);

        monthlySummeryGridPane.getChildren().addAll(totalWeightMonthlyCHN, totalWeightMonthlyLabel, totalMonthlyCHN
                , tableTotalMonthlyLabel, totalInCNYMonthlyLabel, totalInCNYMonthlyCHN);
        monthSummeryPane.getChildren().addAll(monthlySummeryCNY, monthComboBox, monthlySummeryGridPane);



        dataViewBox.getChildren().add(monthSummeryPane);

        Menu fileMenu = new Menu("文件");
        //Menu: File
        MenuItem chooseFileMenuItem = new MenuItem("打开文件");
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("选择需要读取的Excel表格");
        fileMenu.getItems().addAll(chooseFileMenuItem);
        Menu viewMenu = new Menu("视图");
        //Menu: View
        MenuItem cViewMenuItem = new MenuItem("自定义数据范围");
        viewMenu.getItems().addAll(cViewMenuItem);

        menuBar.getMenus().addAll(fileMenu, viewMenu);

        //fileBox
        SimpleStringProperty fileName = new SimpleStringProperty("未选择文件");
        Label info = new Label();
        info.textProperty().bind(fileName);
        Button process = new Button("读取");
        process.setDisable(true);
        process.translateXProperty().bind(primaryStage.widthProperty()
                .subtract(info.widthProperty().add(50)));
        fileBox.getChildren().addAll(info, process);

        //statue bar 2nd


        readFiles();
        monthComboBox.getItems().addAll(catalogTime);


        Menu outputMenu = new Menu("输出");
        MenuItem outputMenuItem = new MenuItem("按月份输出");
        outputMenu.getItems().add(outputMenuItem);
        outputMenuItem.setOnAction(event -> {
            outputDialog();
        });
        menuBar.getMenus().add(outputMenu);

        ObservableList<String> customersList = FXCollections.observableArrayList();
        customersList.addAll(allData.keySet());
        TableColumn customerNameCol = new TableColumn("");
        customersList.addListener(new ListChangeListener<String>() {
            @Override
            public void onChanged(Change<? extends String> c) {
                customersListView.refresh();
            }
        });
        //List
        customersListView.setItems(customersList);
        selectedCustomer.set(customersList.get(0));
        ObservableList<Car> selectedCars = FXCollections.observableArrayList();
        customersListView.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                selectedCustomer.setValue(newValue);
                customerNameCol.setText(newValue);
                selectedCars.clear();
                try {
                    selectedCars.addAll(allData.get(newValue));
                } catch (NullPointerException e) {

                }
                float totalWeight = 0;
                for (Car car : selectedCars) {
                    totalWeight += car.getWeight();
                }
                tableWeightTotal.setValue(totalWeight);

                float total = 0;
                for (Car c : selectedCars) {
                    total += c.getTotal().floatValue();
                }
                tableTotal.setValue(total);
            }
        });
        //table
        selectedTableView.setItems(selectedCars);
        selectedTableView.setEditable(true);
        TableColumn<Car, String> dateCol = new TableColumn<>("日期");
        dateCol.setCellValueFactory(new PropertyValueFactory("date"));
        dateCol.setMinWidth(100);
        TableColumn<Car, String> carNumberCol = new TableColumn<>("车号");
        carNumberCol.setCellValueFactory(new PropertyValueFactory("carNumber"));
        TableColumn<Car, String> brandNameCol = new TableColumn<>("产品名称");
        brandNameCol.setCellValueFactory(new PropertyValueFactory("brandName"));
        brandNameCol.setPrefWidth(200);
        TableColumn<Car, Float> weightCol = new TableColumn<>("重量");
        weightCol.setCellValueFactory(new PropertyValueFactory("weight"));
        TableColumn<Car, String> priceCol = new TableColumn<>("单价");
        priceCol.setCellValueFactory(new PropertyValueFactory("priceString"));
        priceCol.setCellFactory(TextFieldTableCell.forTableColumn());
        priceCol.setOnEditCommit(new EventHandler<TableColumn.CellEditEvent<Car, String>>() {
            @Override
            public void handle(TableColumn.CellEditEvent<Car, String> event) {
                ((Car) event.getTableView().getItems().get(
                        event.getTablePosition().getRow())).setPriceString(event.getNewValue());
                int index = selectedCars.indexOf((Car) event.getTableView().getItems().get(
                        event.getTablePosition().getRow()));
                for (int i = index + 1; i < selectedCars.size(); i++) {
                    Car car = selectedCars.get(i);
                    car.setPriceString(event.getNewValue());
                }
                selectedTableView.refresh();
                float total = 0;
                for (Car c : selectedCars) {
                    total += c.getTotal().floatValue();
                }
                tableTotal.setValue(total);
            }
        });
        TableColumn<Car, Float> totalCol = new TableColumn<>("总价");
        totalCol.setCellValueFactory(new PropertyValueFactory("total"));
        customerNameCol.getColumns().addAll(dateCol, carNumberCol, brandNameCol, weightCol, priceCol, totalCol);
        customerNameCol.setPrefWidth(500);
        selectedTableView.getColumns().addAll(customerNameCol);

        monthComboBox.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                ArrayList<Car> filteredCars = new ArrayList<>();
                filteredCars.addAll(selectedCars);
                filteredCars.removeIf(car -> !car.getDate().substring(0, 7).equals(newValue));
                float totalW = 0;
                float totalP = 0;
                String totalRMB = "";
                for (Car c : filteredCars) {
                    totalW += c.getWeight();
                    totalP += c.getPrice().floatValue() * c.getWeight();
                }
                totalRMB = CNY.number2CNMontrayUnit(new BigDecimal(totalP));
                totalWeightMonthlyLabel.setText(String.valueOf(totalW));
                tableTotalMonthlyLabel.setText(String.valueOf(totalP));
                totalInCNYMonthlyLabel.setText(totalRMB);
            }
        });

        SimpleStringProperty totalInCNY = new SimpleStringProperty("");
        Label customerLable = new Label();
        customerLable.textProperty().bind(selectedCustomer);
        GridPane.setConstraints(customerLable, 1,0);
        Label customerCHN = new Label("客户：");
        GridPane.setConstraints(customerCHN,0,0);
        Label totalWeightCHN = new Label("重量总计（吨）：");
        GridPane.setConstraints(totalWeightCHN,0,1);
        Label totalWeightLabel = new Label();
        totalWeightLabel.textProperty().bind(tableWeightTotal.asString());
        GridPane.setConstraints(totalWeightLabel,1,1);
        Label totalCHN = new Label("金额合计（元）：");
        GridPane.setConstraints(totalCHN,0,2);
        Label tableTotalLabel = new Label();
        tableTotalLabel.textProperty().bind(tableTotal.asString());
        GridPane.setConstraints(tableTotalLabel,1,2);
        Label totalInCNYCHN = new Label("大写人民币：");
        GridPane.setConstraints(totalInCNYCHN, 0,3);
        Label totalInCNYLabel = new Label();
        totalInCNYLabel.textProperty().bind(totalInCNY);
        GridPane.setConstraints(totalInCNYLabel,1,3);

        tableTotal.addListener(new ChangeListener<Number>() {
            @Override
            public void changed(ObservableValue<? extends Number> observable, Number oldValue, Number newValue) {
                totalInCNY.setValue(CNY.number2CNMontrayUnit(new BigDecimal(tableTotal.floatValue())));
            }
        });

        summeryPane.getChildren().addAll(customerCHN, customerLable, totalWeightCHN, totalWeightLabel
            , totalCHN, tableTotalLabel, totalInCNYCHN, totalInCNYLabel);

        chooseFileMenuItem.setOnAction(event -> {
            try {
                workbookFile = fileChooser.showOpenDialog(primaryStage);
            } catch (Exception e) {
                errorReport("未选择文件！");
            }
            openDialog("是否要扫描此文件", event1 -> {
                dataMap.clear();
                dateFileName = "";
                isDateFileNameRegistered = false;
                scan();
                saveFiles();
                System.out.println("+" + dateFileName);
                allData.clear();
                readFiles();
                customersList.setAll(allData.keySet());
                selectedTableView.refresh();
            });

        });
    }

    public static void main(String[] args) throws IOException, InvalidFormatException, OpenXML4JRuntimeException {
        Application.launch(args);
    }
}
