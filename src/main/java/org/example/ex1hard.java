package org.example;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ex1hard {
    public static void main(String[] args) {
        String url = "jdbc:mysql://localhost:3306/bi242";
        String user = "root";
        String password = "2dT#k9H!mQvL7pZ&bR1aWf";

        try (Connection connection = DriverManager.getConnection(url, user, password)) {
            createTableIfNotExists(connection);
            Statement statement = connection.createStatement();

            while (true) {
                System.out.println("Выберите действие:");
                System.out.println("1. Вывести все таблицы из MySQL");
                System.out.println("2. Создать таблицу");
                System.out.println("3. Сложение чисел");
                System.out.println("4. Вычитание чисел");
                System.out.println("5. Умножение чисел");
                System.out.println("6. Деление чисел");
                System.out.println("7. Деление по модулю");
                System.out.println("8. Модуль числа");
                System.out.println("9. Возведение в степень");
                System.out.println("10. Сохранить результаты в Excel");
                System.out.println("0. Выход");

                Scanner scanner = new Scanner(System.in);
                int choice = scanner.nextInt();

                if (choice == 0) break;

                switch (choice) {
                    case 1 -> showTables(statement);
                    case 2 -> createTable(statement);
                    case 3 -> performOperation(connection, "+");
                    case 4 -> performOperation(connection, "-");
                    case 5 -> performOperation(connection, "*");
                    case 6 -> performOperation(connection, "/");
                    case 7 -> performOperation(connection, "%");
                    case 8 -> performUnaryOperation(connection, "abs");
                    case 9 -> performOperation(connection, "^");
                    case 10 -> saveToExcel(statement);
                    default -> System.out.println("Неверный выбор!");
                }
                System.out.println();
            }
        } catch (SQLException | IOException e) {
            e.printStackTrace();
        }
    }

    private static void createTableIfNotExists(Connection connection) throws SQLException {
        String sql = "CREATE TABLE IF NOT EXISTS tables (" +
                "id INT AUTO_INCREMENT PRIMARY KEY, " +
                "operator VARCHAR(20), " +
                "first_operand DOUBLE, " +
                "second_operand DOUBLE, " +
                "result DOUBLE)";
        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(sql);
        }
    }

    private static void showTables(Statement statement) throws SQLException {
        ResultSet resultSet = statement.executeQuery("SHOW TABLES");
        while (resultSet.next()) {
            System.out.println(resultSet.getString(1));
        }
    }

    private static void createTable(Statement statement) throws SQLException {
        statement.executeUpdate("CREATE TABLE IF NOT EXISTS tables (" +
                "id INT AUTO_INCREMENT PRIMARY KEY, " +
                "operator VARCHAR(20), " +
                "first_operand DOUBLE, " +
                "second_operand DOUBLE, " +
                "result DOUBLE)");
    }

    private static void performOperation(Connection connection, String operator) throws SQLException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Введите первое число:");
        double first = scanner.nextDouble();
        System.out.println("Введите второе число:");
        double second = scanner.nextDouble();
        double result = switch (operator) {
            case "+" -> first + second;
            case "-" -> first - second;
            case "*" -> first * second;
            case "/" -> second != 0 ? first / second : Double.NaN;
            case "%" -> second != 0 ? first % second : Double.NaN;
            case "^" -> Math.pow(first, second);
            default -> 0;
        };
        System.out.println("Результат: " + result);
        saveResult(connection, operator, first, second, result);
    }

    private static void performUnaryOperation(Connection connection, String operator) throws SQLException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Введите число:");
        double number = scanner.nextDouble();
        double result = Math.abs(number);
        System.out.println("Модуль числа: " + result);
        saveResult(connection, operator, number, 0, result);
    }

    private static void saveResult(Connection connection, String operator, double first, double second, double result) throws SQLException {
        String sql = "INSERT INTO tables (operator, first_operand, second_operand, result) VALUES (?, ?, ?, ?)";
        try (PreparedStatement preparedStatement = connection.prepareStatement(sql)) {
            preparedStatement.setString(1, operator);
            preparedStatement.setDouble(2, first);
            preparedStatement.setDouble(3, second);
            preparedStatement.setDouble(4, result);
            preparedStatement.executeUpdate();
        }
    }

    private static void saveToExcel(Statement statement) throws SQLException, IOException {
        ResultSet resultSet = statement.executeQuery("SELECT * FROM tables");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Оператор");
        headerRow.createCell(2).setCellValue("Первый операнд");
        headerRow.createCell(3).setCellValue("Второй операнд");
        headerRow.createCell(4).setCellValue("Результат");

        int rowNum = 1;
        while (resultSet.next()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(resultSet.getInt("id"));
            row.createCell(1).setCellValue(resultSet.getString("operator"));
            row.createCell(2).setCellValue(resultSet.getDouble("first_operand"));
            row.createCell(3).setCellValue(resultSet.getDouble("second_operand"));
            row.createCell(4).setCellValue(resultSet.getDouble("result"));
        }

        try (FileOutputStream outputStream = new FileOutputStream("results.xlsx")) {
            workbook.write(outputStream);
        }
        System.out.println("Результаты сохранены в results.xlsx");
    }
}
