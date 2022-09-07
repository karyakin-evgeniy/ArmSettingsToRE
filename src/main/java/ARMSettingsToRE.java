import io.appium.java_client.MobileElement;
import io.appium.java_client.windows.WindowsDriver;
import io.appium.java_client.windows.WindowsElement;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.openqa.selenium.*;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import setting.Setting;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.List;
import java.util.*;
import java.util.concurrent.TimeUnit;

public class ARMSettingsToRE {

    private static final String urlWinApp = "http://127.0.0.1:4723/";
    private static WindowsDriver<WindowsElement> driver = null;
    private static final DesiredCapabilities capForOpenARM = new DesiredCapabilities();
    private static final int timer = 3000;
    private static Actions actions;
    private static int testStatus = 0;
    private static JSONObject properties;
    private static List<Setting> settingTransformList = new LinkedList<>();
    private static List<Setting> automationSettingList = new LinkedList<>();
    private static List<Setting> timeSettingList = new LinkedList<>();
    private static List<Setting> integerSettingList = new LinkedList<>();
    private static List<Setting> programKeyList = new LinkedList<>();
    private static List<Setting> chronophoreSettingList = new LinkedList<>();
    private static List<Setting> allSettings = new LinkedList<>();
    private static List<String> settingsNameNotAddedToMenu = new ArrayList<>();
    private static LinkedHashMap<String, List<Setting>> resultSettingsForExcel = new LinkedHashMap<>();
    private static int countSettingPrograms = 0;

    private static String bfpoDirectory;
    private static String bfpoName = null;

    public static void main(String[] args) throws Exception {
        if (args.length == 1) {
            bfpoName = args[0];
        }
        setUp();
        closeExclels();
        RECreator();

        Thread.sleep(5000);
        Runtime.getRuntime().exit(testStatus);

        try {
            driver.close();
        } catch (Exception ignored) {}
    }

    public static void RECreator() throws Exception {
        openApp(capForOpenARM);
        actions = new Actions(driver);
        List<String> notSettingsNameInMenu = new ArrayList<>();
        notSettingsNameInMenu.add("Положение");
        notSettingsNameInMenu.add("По вертикали");
        notSettingsNameInMenu.add("Строка вниз");

        if(openFileARM(bfpoDirectory, bfpoName)) {
            sleep(3);
            String pathToExcel = properties.getJSONObject("paths").getString("pathForExcel");
            countSettingPrograms = 8;

            findElementByNameAndClick("Коэффициенты трансформации");
            Thread.sleep(timer);
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Коэффициенты трансформации");

            findElementByNameAndClick("Уставки защит и автоматики");
            Thread.sleep(timer);
            for (int i = 1; i < 9; i++) {
                if (checkHaveElementWithName("Значения по программам\n" + i)) {
                    countSettingPrograms = i;
                }
            }
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Уставки защит и автоматики");

            findElementByNameAndClick("Уставки по времени");
            Thread.sleep(timer);
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Уставки по времени");

            findElementByNameAndClick("Целочисленные уставки защит и автоматики");
            Thread.sleep(timer);
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Целочисленные уставки защит и автоматики");

            findElementByNameAndClick("Программные ключи");
            Thread.sleep(timer);
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Программные ключи");

            findElementByNameAndClick("Уставки хронофоры");
            Thread.sleep(timer);
            findElementByNameAndClick("Экспорт в Excel");
            saveFileWithDirectory(pathToExcel, "Уставки хронофоры");


            closeExclels();


//        =================================================================
//        Коэффициенты трансформации

            addSettingsToList(pathToExcel, "Коэффициенты трансформации.xls", settingTransformList, "transform");


//        =================================================================
//        Уставки защит и автоматики

            addSettingsToList(pathToExcel, "Уставки защит и автоматики.xls", automationSettingList, "automatic");

//        =================================================================
//        Уставки по времени

            addSettingsToList(pathToExcel, "Уставки по времени.xls", timeSettingList, "");


//        =================================================================
//        Целочисленные уставки защит и автоматики

            addSettingsToList(pathToExcel, "Целочисленные уставки защит и автоматики.xls", integerSettingList, "integer");

//        =================================================================
//        Программные ключи

            addSettingsToList(pathToExcel, "Программные ключи.xls", programKeyList, "bool");


            sleep(10);
            System.out.println("Количество уставок по разделам");
            System.out.println("Коэффициенты трансформации " + settingTransformList.size());
            System.out.println("Уставки защит и автоматики " + automationSettingList.size());
            System.out.println("Уставки по времени " + timeSettingList.size());
            System.out.println("Целочисленные уставки защит и автоматики " + integerSettingList.size());
            System.out.println("Программные ключи " + programKeyList.size());
            System.out.println("Уставки хронофоры " + chronophoreSettingList.size());
            System.out.println("Все уставки " + allSettings.size());



//                ===========================================================================
//                Получение всех уставок из списка Редактор меню монитора


            findElementByNameAndClick("Редактор меню монитора");
            sleep(10);


            WindowsElement settingsElement = driver.findElementByName("Уставки, конфигурация");
            settingsElement.click();
//                Получение всех childElement для раздела "Уставки, конфигурация"
            List<MobileElement> elements = settingsElement.findElementsByXPath("//*");
            LinkedHashMap<String, List<String>>  monitorMenu = new LinkedHashMap<>();
//                System.out.println(elements.size());
            List<String> lastSectionElements = new LinkedList<>();
            HashSet<String> allSettingsName = new HashSet<>(settingsNameNotAddedToMenu);
            HashSet<String> allSettingsInMenu = new HashSet<>();
            for (MobileElement element : elements) {
//                    System.out.println(element.getAttribute("Name"));
                String sectionName = element.getAttribute("Name");
                try {
                    sleep(1);
                    element.click();
                    sleep(1);
                    element.click();
                    sleep(1);
                    element.click();
                    sleep(1);
                } catch (WebDriverException e) {
                    findElementByNameAndClick("Прокрутить вниз");
                }
                List<MobileElement> settingsInSection = driver.findElementByClassName("SysListView32").findElementsByXPath("//*");
                List<String> settingsNameInSection = new LinkedList<>();
                if (settingsInSection.size() > 1 && lastSectionElements.contains(settingsInSection.get(1).getAttribute("Name"))) {
                    element.click();
                    settingsInSection = driver.findElementByClassName("SysListView32").findElementsByXPath("//*");
                }
                lastSectionElements = new LinkedList<>();
                for (int i = 1; i < settingsInSection.size(); i+=2) {
//                        System.out.println(settingsInSection.get(i).getAttribute("Name"));
                    String name = settingsInSection.get(i).getAttribute("Name");
                    settingsNameInSection.add(name);
                    settingsNameNotAddedToMenu.remove(name);
                    allSettingsInMenu.add(name);
                    lastSectionElements.add(name);
                }
                monitorMenu.put(sectionName, settingsNameInSection);
//                    System.out.println(settingsNameInSection.size());


            }
            sleep(20);
            monitorMenu.remove("Уставки, конфигурация");

            for (String sectionName : monitorMenu.keySet()) {
                List<String> settingsNameInSection = monitorMenu.get(sectionName);
                List<Setting> settings = new ArrayList<>();
                for (String settingName : settingsNameInSection) {
                    allSettings.forEach(setting -> {
                        if (setting.getName().equals(settingName)) {
                            settings.add(setting);
                        }
                    });
                }
                resultSettingsForExcel.put(sectionName, settings);
            }

            System.out.println("Уставки, которые не попали в список \"Редактор меню блока\":");
            settingsNameNotAddedToMenu.forEach(System.out::println);
            List<Setting> notMenuSettingsForResult = new ArrayList<>();
            for (String settingName : settingsNameNotAddedToMenu) {
                for (Setting setting : allSettings) {
                    if (settingName.equals(setting.getName())) {
                        notMenuSettingsForResult.add(setting);
                        break;
                    }
                }
            }
            resultSettingsForExcel.put("Уставки не попавшие в \"Редактор меню монитора\"", notMenuSettingsForResult);

            System.out.println("Уставки, которые есть в меню монитора, но отсутствуют в уставках:");
            allSettingsInMenu.removeAll(allSettingsName);
            allSettingsInMenu.removeAll(notSettingsNameInMenu);
            allSettingsInMenu.forEach(System.out::println);

//                ===========================================================================
//                Редактор меню блока


            findElementByNameAndClick("Редактор меню блока");
            sleep(10);


            settingsElement = driver.findElementByName("Уставки, конфигурация");
            settingsElement.click();
//                Получение всех childElement для раздела "Уставки, конфигурация"
            elements = settingsElement.findElementsByXPath("//*");
            LinkedHashMap<String, List<String>>  blockMenu = new LinkedHashMap<>();
            HashSet<String> sectionsName = new HashSet<>();
            for (MobileElement element : elements) {
                sectionsName.add(element.getAttribute("Name"));
            }
            for (MobileElement element : elements) {
//                    System.out.println(element.getAttribute("Name"));
                String sectionName = element.getAttribute("Name");
                try {
                    sleep(1);
                    element.click();
                    sleep(1);
                    element.click();
                    sleep(1);
                    element.click();
                    sleep(1);
                } catch (WebDriverException e) {
                    findElementByNameAndClick("Прокрутить вниз");
                }
                List<MobileElement> settingsInSection = driver.findElementByClassName("SysListView32").findElementsByXPath("//*");
                List<String> settingsNameInSection = new LinkedList<>();
                if (settingsInSection.size() > 1 && lastSectionElements.contains(settingsInSection.get(1).getAttribute("Name"))) {
                    element.click();
                    settingsInSection = driver.findElementByClassName("SysListView32").findElementsByXPath("//*");
                }
                lastSectionElements = new LinkedList<>();
                for (int i = 1; i < settingsInSection.size(); i+=2) {
                    String name = settingsInSection.get(i).getAttribute("Name");
                    if (sectionsName.contains(name)) {
                        i = 1000;
                    } else {
//                            System.out.println(name);
                        settingsNameInSection.add(name);
                        settingsNameNotAddedToMenu.remove(name);
                        allSettingsInMenu.add(name);
                        lastSectionElements.add(name);
                    }
                }
                blockMenu.put(sectionName, settingsNameInSection);
//                    System.out.println(settingsNameInSection.size());


            }
            sleep(20);
            blockMenu.remove("Уставки, конфигурация");


            for (String sectionName : monitorMenu.keySet()) {
                List<String> monitorMenuList = monitorMenu.get(sectionName);
                List<String> blockMenuList = blockMenu.get(sectionName);

                if (blockMenuList == null && (sectionName.startsWith("Ускорение") || sectionName.startsWith("Дополнительные") || sectionName.startsWith("Контроль") || sectionName.startsWith("Защита"))) {
                    blockMenuList = blockMenu.get(sectionName.replace("Ускорение", "Уск.")
                            .replace("Дополнительные", "Доп.")
                            .replace("Контроль напряжений", "КН")
                            .replace("Контроль синхронизма", "КС")
                            .replace("Защита", "Защ."));
                }

                if (blockMenuList == null || monitorMenuList == null) {
                    System.out.println("В разделе " + sectionName + " меню блока = null");
                } else {
                    blockMenuList.removeAll(notSettingsNameInMenu);
                    monitorMenuList.removeAll(notSettingsNameInMenu);
                    if (blockMenuList.size() == monitorMenuList.size()) {
                        blockMenuList.forEach(blockSettingNameOriginal -> {
                            boolean checkHaveSettingInBlock = false;
                            for (String settingName : monitorMenuList) {
                                settingName = settingName.replaceAll(" ", "").replace("(", "").replace(")", "");
                                String blockSettingName = blockSettingNameOriginal.replaceAll(" ", "").replace("(", "").replace(")", "");
                                if (blockSettingName.startsWith(settingName)
                                        && blockSettingName.length() >= settingName.length()
                                        && blockSettingName.length() < settingName.length() + 8) {
                                    checkHaveSettingInBlock = true;
                                    break;
                                } else if (settingName.length() > 10 && blockSettingName.startsWith(settingName.substring(0, 8))) {
                                    checkHaveSettingInBlock = true;
                                    break;
                                }
                            }
                            if (!checkHaveSettingInBlock) {
                                System.out.println("Уставка " + blockSettingNameOriginal + " отсутствует в меню монитора");
                            }
                        });
                    } else {
                        List<String> blockMenuListMinusMonitor = new ArrayList<>(blockMenuList);
                        List<String> monitorMenuListMinusBlock = new ArrayList<>(monitorMenuList);
                        blockMenuListMinusMonitor.removeAll(monitorMenuList);
                        System.out.println("Уставки которые есть в меню блока, но нет в меню монитора:");
                        blockMenuListMinusMonitor.forEach(System.out::println);
                        System.out.println("Уставки которые есть в меню монитора, но нет в меню блока:");
                        monitorMenuListMinusBlock.removeAll(blockMenuList);
                        monitorMenuList.forEach(System.out::println);
                        System.out.println("Размеры раздела " + sectionName + " отличаются. Монитор меню: " + monitorMenuList.size()
                                + ". Монитор блока: " + blockMenuList.size());
                    }
                }
            }


            closeARM();


//        ============================================================
//        ============================================================
//        Запись в новый файл


            HSSFWorkbook writeWorkbook = new HSSFWorkbook();
            HSSFSheet writeSheet = writeWorkbook.createSheet();
            int rowNum = 0;
            Cell cell;
            Row rowForWrite;

            HSSFCellStyle style = createStileForTitle(writeWorkbook);
            HSSFCellStyle styleForValue = createStileForValue(writeWorkbook);
            HSSFCellStyle styleForName = createStileForName(writeWorkbook);


//        ==================================================
//        Название столбцов


            for (String sectionName : resultSettingsForExcel.keySet()) {
                rowForWrite = writeSheet.createRow(rowNum);

                writeCell(0, rowForWrite, sectionName, style);
                writeSheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 4));
                rowNum++;
                List<Setting> settingList = resultSettingsForExcel.get(sectionName);

                if (settingList.size() != 0) {
                    rowNaming(rowNum, writeSheet, style);
                    rowNum++;
                }

                for (Setting setting : settingList) {
                    Row row = writeSheet.createRow(rowNum);

                    writeCell(0, row, setting.getSetting(), styleForName);

                    writeCell(1, row, setting.getName(), styleForName);

                    writeCell(2, row, setting.getValue(), styleForValue);

                    if (!setting.getType().equals("bool")) {

                        writeCell(3, row, setting.getStartRange() + " - " + setting.getFinishRange(), styleForValue);

                        writeCell(4, row, setting.getStep(), styleForValue);

                    } else {
                        writeCell(3, row, "ключ", styleForValue);

                        writeCell(4, row, "-", styleForValue);

                    }
                    rowNum++;

                }
            }


            File resultFile = new File(properties.getJSONObject("paths").getString("pathForResultExcel")
                    + "\\allSettingsForProject-" + bfpoName + ".xls");
            resultFile.getParentFile().mkdirs();

            FileOutputStream outFile = new FileOutputStream(resultFile);
            writeWorkbook.write(outFile);
            System.out.println("Created file: " + resultFile.getAbsolutePath());

        }
    }

    private static void rowNaming(int rowNum, HSSFSheet writeSheet, HSSFCellStyle style) {
        Row rowForWrite = writeSheet.createRow(rowNum);


        writeCell(0, rowForWrite, "Уставка",style);

        writeCell(1, rowForWrite, "Обозначение",style);

        writeCell(2, rowForWrite, "Заводская установка",style);

        writeCell(3, rowForWrite, "Диапазон значений",style);

        writeCell(4, rowForWrite, "Дискретность",style);

    }

    private static void addSettingsToList(String pathToExcel, String fileName, List<Setting> settingList, String type) throws IOException {
        FileInputStream inputStream = new FileInputStream(pathToExcel + "\\" + fileName);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next();

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            Setting setting = new Setting();
            Iterator<Cell> cellIterator = row.cellIterator();
            int i = 0;
            try {
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    CellType cellType = cell.getCellType();

//                    printCell(cellType, cell);

                    if (i < 2 || type.equals("transform")) {
                        if (addSettingsParam(setting, i, cell, type)) {
                            settingList.add(setting);
                            settingsNameNotAddedToMenu.add(setting.getName());
                        }
                    } else if (i <= countSettingPrograms) {
                        if (addSettingsParam(setting, 2, cell, type)) {
                            settingList.add(setting);
                        }
                    } else {
                        if (addSettingsParam(setting, i - countSettingPrograms + 2, cell, type)) {
                            settingList.add(setting);
                            settingsNameNotAddedToMenu.add(setting.getName());
                        }
                    }
                    i++;
                }
            } catch (NoSuchElementException ignored) {}
//            System.out.println("");
        }
        allSettings.addAll(settingList);
    }

    private static HSSFCellStyle createStileForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight((short) 240);
        font.setFontName("Times New Roman");
        HSSFCellStyle style = workbook.createCellStyle();
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        style.setFillBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());

        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static HSSFCellStyle createStileForValue(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontHeight((short) 240);
        font.setFontName("Times New Roman");
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static HSSFCellStyle createStileForName(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontHeight((short) 240);
        font.setFontName("Times New Roman");
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return style;
    }

    private static void writeCell(int i, Row row, String cellValue, HSSFCellStyle styleForValue) {
        Cell cell = row.createCell(i, CellType.STRING);
        cell.setCellValue(cellValue);

        cell.setCellStyle(styleForValue);

    }

    public static boolean addSettingsParam(Setting setting, int i, Cell cell, String type) {
        boolean addToList = false;

        switch (type) {
            case "transform":
                switch (i) {
                    case 0:
                        setting.setName(cell.getStringCellValue());

                        break;
                    case 2:
                        setting.setValue(cell.getStringCellValue());
                        break;
                    case 3:
                        setting.setStartRange(cell.getStringCellValue());
                        break;

                    case 4:
                        setting.setFinishRange(cell.getStringCellValue());
                        break;

                    case 5:
                        setting.setStep(cell.getStringCellValue());
                        break;

                    case 6:
                        addToList = checkAddASU(cell, setting);

                        break;
                    case 7:
                        setting.setSetting(cell.getStringCellValue());
                }
                break;
            case "bool":
                switch (i) {
                    case 0:
                        setting.setName(cell.getStringCellValue());
                        break;
                    case 1:
                        setting.setValue(cell.getStringCellValue());
                        break;
                    case 2:
                        if (checkEqualsValue(cell.getStringCellValue(), setting)) {
                            setting.setRight(false);
                            setting.getAllValues().add(setting.getValue());
                            setting.getAllValues().add(cell.getStringCellValue());
                        }
                        break;
                    case 3:
                        if (!setting.isRight()) {
                            System.out.println(setting.getName() + " уставки значений по программам не равны");
                            setting.getAllValues().forEach(System.out::println);
                        }
                        setting.setType("bool");
                        addToList = checkAddASU(cell, setting);

                        break;
                    case 5:
                        setting.setSetting(cell.getStringCellValue());
                }
                break;
            case "integer":
                switch (i) {
                    case 0:
                        setting.setName(cell.getStringCellValue());
                        break;
                    case 1:
                        setting.setValue(cell.getStringCellValue());
                        break;
                    case 2:
                        if (checkEqualsValue(cell.getStringCellValue(), setting)) {
                            setting.setRight(false);
                            setting.getAllValues().add(setting.getValue());
                            setting.getAllValues().add(cell.getStringCellValue());
                        }
                        break;
                    case 3:
                        setting.setStartRange(cell.getStringCellValue());
                        break;
                    case 4:
                        setting.setFinishRange(cell.getStringCellValue());
                        break;
                    case 5:
                        if (!setting.isRight()) {
                            System.out.println(setting.getName() + " уставки значений по программам не равны");
                            setting.getAllValues().forEach(System.out::println);
                        }
                        addToList = checkAddASU(cell, setting);
                        setting.setStep("1");

                        break;
                    case 7:
                        setting.setSetting(cell.getStringCellValue());
                }
                break;
            case "automatic":
                switch (i) {
                    case 0:
                        setting.setName(cell.getStringCellValue());
                        break;
                    case 1:
                        setting.setValue(cell.getStringCellValue());
                        break;
                    case 2:
                        if (checkEqualsValue(cell.getStringCellValue(), setting)) {
                            setting.setRight(false);
                            setting.getAllValues().add(setting.getValue());
                            setting.getAllValues().add(cell.getStringCellValue());
                        }
                        break;
                    case 3:
                        setting.setStartRange(cell.getStringCellValue());
                        break;
                    case 4:
                        setting.setFinishRange(cell.getStringCellValue());
                        break;
                    case 5:
                        setting.setStep(cell.getStringCellValue());
                        break;
                    case 6:
                        if (!setting.isRight()) {
                            System.out.println(setting.getName() + " уставки значений по программам не равны");
                            setting.getAllValues().forEach(System.out::println);
                        }
                        addToList = checkAddASU(cell, setting);

                        break;
                    case 9:
                        setting.setSetting(cell.getStringCellValue());
                }
                break;
            default:
                switch (i) {
                    case 0:
                        setting.setName(cell.getStringCellValue());
                        break;
                    case 1:
                        setting.setValue(cell.getStringCellValue());
                        break;
                    case 2:
                        if (checkEqualsValue(cell.getStringCellValue(), setting)) {
                            setting.setRight(false);
                            setting.getAllValues().add(setting.getValue());
                            setting.getAllValues().add(cell.getStringCellValue());
                        }
                        break;
                    case 3:
                        setting.setStartRange(cell.getStringCellValue());
                        break;
                    case 4:
                        setting.setFinishRange(cell.getStringCellValue());
                        break;
                    case 5:
                        setting.setStep(cell.getStringCellValue());
                        break;
                    case 6:
                        if (!setting.isRight()) {
                            System.out.println(setting.getName() + " уставки значений по программам не равны");
                            setting.getAllValues().forEach(System.out::println);
                        }
                        addToList = checkAddASU(cell, setting);
                        break;
                    case 8:
                        setting.setSetting(cell.getStringCellValue());
                }
        }
        return addToList;
    }


    private static boolean checkAddASU(Cell cell, Setting setting) {
        boolean addToList = false;
        if (cell.getStringCellValue().equals("+")) {
            setting.setAddASU(true);
            addToList = true;
        }
        if (!properties.getBoolean("checkAddToASU")) {
            addToList = true;
        }
        return addToList;
    }
    private static boolean checkEqualsValue(String readValue, Setting setting) {
        return !readValue.equals(setting.getValue()) && !readValue.equals(setting.getValue().replace(",", "."));
    }


    public static void findElementByNameAndClick(String name) {
        driver.findElementByName(name).click();
    }

    public static boolean openFileARM(String directory, String fileName) throws InterruptedException {
        driver.findElementByName("Стандартный").findElementByName("Открыть").click();

        return openFileWithDirectory(directory, fileName);
    }

    public static boolean openPMKFile(String directory, String fileName) throws InterruptedException {
        findElementByNameAndClick("Открыть");
        driver.getKeyboard().sendKeys(Keys.CONTROL + "o" + Keys.CONTROL);
        return openFileWithDirectory(directory, fileName);
    }

    public static boolean openFileWithDirectory(String directory, String fileName) throws InterruptedException {


        driver.findElementByName("Все папки").click();
        Thread.sleep(timer);
        WindowsElement directoryField =  driver.findElementByName("Адрес");
        String startDirectory = directoryField.getAttribute("Value.Value");


        if (!directory.equals(startDirectory)) {
            sendKeys(directory);
        }

        return openFile(fileName);


    }

    public static boolean openFile(String name) throws InterruptedException {
        try {
            addToCopyBuffer(name);
            driver.findElementByClassName("ComboBox").findElementByName("Имя файла:").sendKeys(Keys.CONTROL + "v" + Keys.ENTER);
            Thread.sleep(10000);
        } catch (NoSuchElementException e) {


            return false;

        }
        return true;
    }

    public static boolean saveFileWithDirectory(String directory, String fileName) throws InterruptedException {


        driver.findElementByName("Все папки").click();
        Thread.sleep(timer);
        WindowsElement directoryField =  driver.findElementByName("Адрес");
        String startDirectory = directoryField.getAttribute("Value.Value");


        if (!directory.equals(startDirectory)) {
            sendKeys(directory);
        }

        return saveFile(fileName);


    }

    public static boolean saveFile(String name) throws InterruptedException {
        try {
            addToCopyBuffer(name);
            driver.findElementByClassName("Edit").findElementByName("Имя файла:").sendKeys(Keys.CONTROL + "v" + Keys.ENTER);
            if (checkHaveElementWithName("Подтвердить сохранение в виде")) {
                findElementByNameAndClick("Да");
            }
            Thread.sleep(10000);
        } catch (NoSuchElementException e) {


            return false;

        }
        return true;
    }




    public static boolean checkHaveElementWithName(String name) {
        try {
//            System.out.println("Поиск элемента с именем - " + name);
            driver.findElementByName(name);
            return true;
        } catch (NoSuchElementException | NoSuchWindowException e) {
            System.out.println("Элемент с именем - " + name + " не найден");
            return false;
        }
    }
    private static void expandApp() {
        try {
            driver.findElementByName("Развернуть").click();
        } catch (NoSuchElementException | NoClassDefFoundError ignored) {}
    }

    private static void setUp() throws IOException {
        String textProperties = new String(Files.readAllBytes(Paths.get(".\\jsonProperties\\properties.json")), StandardCharsets.UTF_8);
        properties = new JSONObject(textProperties);
        JSONObject paths = properties.getJSONObject("paths");
        String armPath = paths.getString("pathToArm");
        capForOpenARM.setCapability("app", armPath);
        JSONObject bfpoProperties = properties.getJSONObject("bfpo");
        bfpoDirectory = bfpoProperties.getString("bfpoDirectory");
        if (bfpoName == null) {
            bfpoName = bfpoProperties.getString("bfpoName");
        }
    }

    public static void openApp(DesiredCapabilities appCap) {
        try {
            driver = new WindowsDriver<WindowsElement>(new URL(urlWinApp), appCap);
            driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
            expandApp();
        } catch (MalformedURLException e) {
            e.printStackTrace();
        }

    }

    public static void sleep(double seconds) throws InterruptedException {
        Thread.sleep((long) (seconds * 1000));
    }

    public static void addToCopyBuffer(String copy) {
        StringSelection stringSelection = new StringSelection(copy);
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        clipboard.setContents(stringSelection, null);
    }

    public static void sendKeys(String name) {
        addToCopyBuffer(name);
        driver.getKeyboard().sendKeys(Keys.CONTROL + "v" + Keys.CONTROL + Keys.ENTER);
    }
    public static void closeExclels() {
        try {
            Runtime.getRuntime().exec("taskkill /IM EXCEL.EXE /f");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void closeARM() {
        findElementByNameAndClick("Закрыть");
        findElementByNameAndClick("Нет");
    }
}
