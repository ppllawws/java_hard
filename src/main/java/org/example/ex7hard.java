package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.File;
import java.sql.*;
import java.util.Scanner;
import java.util.Arrays;

public class ex7hard {
    private static final String URL = "jdbc:mysql://localhost:3306/my_database";
    private static final String USER = "root";
    private static final String PASSWORD = "2dT#k9H!mQvL7pZ&bR1aWf";
    private static final int SIZE = 35;
    private static int[] matrix = new int[SIZE];
    private static String tableName = "";

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        while (true) {
            System.out.println("\nВыберите действие:");
            System.out.println("1. Вывести все таблицы из MySQL");
            System.out.println("2. Создать таблицу");
            System.out.println("3. Ввести одномерный массив и сохранить в MySQL");
            System.out.println("4. Отсортировать массив и сохранить в MySQL и вывести");
            System.out.println("5. Сохранить результат в Excel");
            System.out.println("0. Выход");
            System.out.print("Ваш выбор: ");

            int choice;
            if (scanner.hasNextInt()) {
                choice = scanner.nextInt();
                scanner.nextLine();
            } else {
                System.out.println("Ошибка ввода. Введите число от 0 до 5.");
                scanner.next();
                continue;
            }

            switch (choice) {
                case 1 -> listTables();
                case 2 -> createTable(scanner);
                case 3 -> saveArrayToDatabase(scanner);
                case 4 -> sortArray();
                case 5 -> exportToExcel();
                case 0 -> {
                    System.out.println("Выход из программы.");
                    scanner.close();
                    return;
                }
                default -> System.out.println("Ошибка ввода. Введите число от 0 до 5.");
            }
        }
    }

    private static void listTables() {
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery("SHOW TABLES")) {

            System.out.println("Список таблиц в базе данных:");
            while (rs.next()) {
                System.out.println(rs.getString(1));
            }
        } catch (SQLException e) {
            System.out.println("Ошибка при получении списка таблиц: " + e.getMessage());
        }
    }

    private static void createTable(Scanner scanner) {
        System.out.print("Введите имя таблицы: ");
        tableName = scanner.nextLine();

        String sql = "CREATE TABLE IF NOT EXISTS `" + tableName + "` (" +
                "id INT AUTO_INCREMENT PRIMARY KEY," +
                "index_value INT," +
                "array_value INT)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement()) {
            stmt.execute(sql);
            System.out.println("Таблица '" + tableName + "' успешно создана или уже существует.");
        } catch (SQLException e) {
            System.out.println("Ошибка при создании таблицы: " + e.getMessage());
        }
    }

    private static void saveArrayToDatabase(Scanner scanner) {
        System.out.println("Введите " + SIZE + " целых чисел:");
        for (int i = 0; i < SIZE; i++) {
            while (true) {
                System.out.print("Элемент " + (i + 1) + ": ");
                if (scanner.hasNextInt()) {
                    matrix[i] = scanner.nextInt();
                    break;
                } else {
                    System.out.println("Ошибка: Введите целое число!");
                    scanner.next(); // Очистка некорректного ввода
                }
            }
        }
        saveToDatabase("original", matrix);
    }

    private static void sortArray() {
        Arrays.sort(matrix);
        saveToDatabase("sorted", matrix);
        System.out.println("Отсортированный массив: " + Arrays.toString(matrix));
    }

    private static void saveToDatabase(String type, int[] array) {
        String sql = "INSERT INTO `" + tableName + "` (index_value, array_value) VALUES (?, ?)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            for (int i = 0; i < SIZE; i++) {
                pstmt.setInt(1, i);
                pstmt.setInt(2, array[i]);
                pstmt.addBatch();
            }
            pstmt.executeBatch();
            System.out.println("Массив ('" + type + "') сохранен в таблицу '" + tableName + "' в базе данных.");
        } catch (SQLException e) {
            System.out.println("Ошибка при сохранении массива: " + e.getMessage());
        }
    }

    private static void exportToExcel() {
        String filePath = "array_data.xlsx";
        File file = new File(filePath);

        try (Workbook workbook = file.exists() ? new XSSFWorkbook(new FileInputStream(file)) : new XSSFWorkbook()) {
            // Проверяем, существует ли лист с именем "Array Data"
            Sheet sheet = workbook.getSheet("Array Data");
            if (sheet == null) {
                sheet = workbook.createSheet("Array Data");
            }

            int rowStart = file.exists() ? sheet.getPhysicalNumberOfRows() : 0;

            for (int i = rowStart; i < SIZE + rowStart; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
                row.createCell(1).setCellValue(matrix[i - rowStart]);
            }

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Результаты сохранены в " + filePath);
        } catch (IOException e) {
            System.out.println("Ошибка при сохранении в Excel: " + e.getMessage());
        }
    }
}
