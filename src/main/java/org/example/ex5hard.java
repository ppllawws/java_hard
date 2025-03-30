package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.InputMismatchException;
import java.util.Scanner;

public class ex5hard {
    private static final String URL = "jdbc:mysql://localhost:3306/my_database";
    private static final String USER = "root";
    private static final String PASSWORD = "2dT#k9H!mQvL7pZ&bR1aWf";
    private static final String FILE_PATH = "results.xlsx";
    private static final int MIN_LENGTH = 50;

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("\nВыберите действие:");
            System.out.println("1. Вывести все таблицы из MySQL");
            System.out.println("2. Создать таблицу");
            System.out.println("3. Изменить порядок символов строки на обратный и сохранить");
            System.out.println("4. Добавить одну строку в другую и сохранить");
            System.out.println("5. Экспортировать данные в Excel");
            System.out.println("0. Выход");

            int choice = getValidChoice(scanner);

            switch (choice) {
                case 1 -> listTables();
                case 2 -> createTable(scanner);
                case 3 -> reverseAndSaveString(scanner);
                case 4 -> concatenateAndSaveStrings(scanner);
                case 5 -> exportToExcel();
                case 0 -> {
                    System.out.println("Выход");
                    return;
                }
            }
        }
    }

    private static int getValidChoice(Scanner scanner) {
        int choice;
        while (true) {
            try {
                choice = scanner.nextInt();
                scanner.nextLine();

                if (choice >= 0 && choice <= 5) {
                    return choice;
                } else {
                    System.out.println("Ошибка! Введите число от 0 до 5.");
                }
            } catch (InputMismatchException e) {
                System.out.println("Ошибка! Введите число от 0 до 5.");
                scanner.nextLine();
            }
        }
    }

    private static void listTables() {
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery("SHOW TABLES")) {
            System.out.println("Список таблиц:");
            while (rs.next()) {
                System.out.println(rs.getString(1));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private static void createTable(Scanner scanner) {
        System.out.print("Введите имя таблицы (не менее 3 символов, не начинается с числа): ");
        String tableName = scanner.nextLine();

        String createTableSQL = "CREATE TABLE IF NOT EXISTS `" + tableName + "` (id INT AUTO_INCREMENT PRIMARY KEY, text VARCHAR(255))";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement()) {
            stmt.executeUpdate(createTableSQL);
            System.out.println("Таблица '" + tableName + "' создана.");
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private static void reverseAndSaveString(Scanner scanner) {
        String input;
        do {
            System.out.print("Введите строку (не менее 50 символов): ");
            input = scanner.nextLine();
            if (input.length() < MIN_LENGTH) {
                System.out.println("Ошибка! Введенная строка слишком короткая. Попробуйте снова.");
            }
        } while (input.length() < MIN_LENGTH);

        String reversed = new StringBuilder(input).reverse().toString();
        saveStringToDatabase(reversed);
        System.out.println("Перевернутая строка: " + reversed);
    }

    private static void concatenateAndSaveStrings(Scanner scanner) {
        String str1, str2;
        do {
            System.out.print("Введите первую строку (не менее 50 символов): ");
            str1 = scanner.nextLine();
            if (str1.length() < MIN_LENGTH) {
                System.out.println("Ошибка! Введенная строка слишком короткая. Попробуйте снова.");
            }
        } while (str1.length() < MIN_LENGTH);

        do {
            System.out.print("Введите вторую строку (не менее 50 символов): ");
            str2 = scanner.nextLine();
            if (str2.length() < MIN_LENGTH) {
                System.out.println("Ошибка! Введенная строка слишком короткая. Попробуйте снова.");
            }
        } while (str2.length() < MIN_LENGTH);

        String combined = str1 + str2;
        saveStringToDatabase(combined);
        System.out.println("Объединенная строка: " + combined);
    }

    private static void saveStringToDatabase(String text) {
        String insertSQL = "INSERT INTO strings (text) VALUES (?)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             PreparedStatement pstmt = conn.prepareStatement(insertSQL)) {
            pstmt.setString(1, text);
            pstmt.executeUpdate();
            System.out.println("Строка сохранена в базе данных.");
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private static void exportToExcel() {
        String query = "SELECT * FROM strings";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(query);
             Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Strings Data");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("Text");

            int rowNum = 1;
            while (rs.next()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(rs.getInt("id"));
                row.createCell(1).setCellValue(rs.getString("text"));
            }

            try (FileOutputStream outputStream = new FileOutputStream(FILE_PATH)) {
                workbook.write(outputStream);
            }
            System.out.println("Результаты сохранены в results.xlsx");

        } catch (SQLException | IOException e) {
            e.printStackTrace();
        }
    }
}
