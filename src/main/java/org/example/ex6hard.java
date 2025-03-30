package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.InputMismatchException;
import java.util.Scanner;

public class ex6hard {
    private static final String URL = "jdbc:mysql://localhost:3306/my_database";
    private static final String USER = "root";
    private static final String PASSWORD = "2dT#k9H!mQvL7pZ&bR1aWf";
    private static final int SIZE = 7;
    private static int[][] matrix1 = new int[SIZE][SIZE];
    private static int[][] matrix2 = new int[SIZE][SIZE];
    private static int[][] resultMatrix = new int[SIZE][SIZE];
    private static String tableName = "";

    public static void main(String[] args) {
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD)) {
            System.out.println("Подключение успешно!");
        } catch (SQLException e) {
            System.out.println("Ошибка подключения: " + e.getMessage());
            return;
        }

        Scanner scanner = new Scanner(System.in);
        while (true) {
            System.out.println("\nВыберите действие:");
            System.out.println("1. Вывести все таблицы из MySQL");
            System.out.println("2. Создать таблицу");
            System.out.println("3. Ввести две матрицы с клавиатуры и сохранить в MySQL");
            System.out.println("4. Перемножить матрицы, сохранить в MySQL и вывести");
            System.out.println("5. Сохранить результат в Excel");
            System.out.println("0. Выход");
            System.out.print("Ваш выбор: ");

            int choice;
            try {
                choice = scanner.nextInt();
            } catch (InputMismatchException e) {
                System.out.println("Ошибка ввода. Введите число от 0 до 5.");
                scanner.nextLine();
                continue;
            }

            switch (choice) {
                case 1 -> listTables();
                case 2 -> createTable(scanner);
                case 3 -> saveMatrixToDatabase(scanner);
                case 4 -> multiplyMatrices();
                case 5 -> exportToExcel();
                case 0 -> {
                    System.out.println("Выход из программы.");
                    scanner.close();
                    return;
                }
                default -> System.out.println("Неверный выбор. Попробуйте снова.");
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
        scanner.nextLine();
        System.out.print("Введите имя таблицы (не менее 3 символов, не начинается с числа): ");
        tableName = scanner.nextLine();

        if (tableName.length() < 3 || Character.isDigit(tableName.charAt(0))) {
            System.out.println("Ошибка: имя таблицы должно содержать минимум 3 символа и не начинаться с цифры!");
            return;
        }

        String createTableSQL = "CREATE TABLE IF NOT EXISTS `" + tableName + "` (id INT AUTO_INCREMENT PRIMARY KEY, text VARCHAR(255))";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement()) {
            stmt.executeUpdate(createTableSQL);
            System.out.println("Таблица '" + tableName + "' создана.");
        } catch (SQLException e) {
            System.out.println("Ошибка при создании таблицы: " + e.getMessage());
        }
    }

    private static void saveMatrixToDatabase(Scanner scanner) {
        System.out.println("Введите элементы первой матрицы:");
        fillMatrix(scanner, matrix1);
        System.out.println("Введите элементы второй матрицы:");
        fillMatrix(scanner, matrix2);

        saveMatrix("matrix1", matrix1);
        saveMatrix("matrix2", matrix2);
    }

    private static void fillMatrix(Scanner scanner, int[][] matrix) {
        for (int i = 0; i < SIZE; i++) {
            for (int j = 0; j < SIZE; j++) {
                while (true) {
                    System.out.print("Введите элемент [" + i + "][" + j + "]: ");
                    if (scanner.hasNextInt()) {
                        matrix[i][j] = scanner.nextInt();
                        break;
                    } else {
                        System.out.println("Ошибка: Введите целое число!");
                        scanner.next();
                    }
                }
            }
        }
    }

    private static void saveMatrix(String name, int[][] matrix) {
        String sql = "INSERT INTO `" + tableName + "` (matrix_name, row_index, col_index, value) VALUES (?, ?, ?, ?)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            for (int i = 0; i < SIZE; i++) {
                for (int j = 0; j < SIZE; j++) {
                    pstmt.setString(1, name);
                    pstmt.setInt(2, i);
                    pstmt.setInt(3, j);
                    pstmt.setInt(4, matrix[i][j]);
                    pstmt.addBatch();
                }
            }
            pstmt.executeBatch();
            System.out.println("Матрица " + name + " сохранена в таблицу '" + tableName + "'.");
        } catch (SQLException e) {
            System.out.println("Ошибка при сохранении матрицы: " + e.getMessage());
        }
    }

    private static void multiplyMatrices() {
        for (int i = 0; i < SIZE; i++) {
            for (int j = 0; j < SIZE; j++) {
                for (int k = 0; k < SIZE; k++) {
                    resultMatrix[i][j] += matrix1[i][k] * matrix2[k][j];
                }
            }
        }

        saveMatrix("result_matrix", resultMatrix);
        printMatrix(resultMatrix, "Результат умножения:");
    }

    private static void printMatrix(int[][] matrix, String title) {
        System.out.println(title);
        for (int[] row : matrix) {
            for (int value : row) {
                System.out.print(value + "\t");
            }
            System.out.println();
        }
    }

    private static void exportToExcel() {
        String filePath = "resultmatrix.xlsx";
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(filePath)) {

            Sheet sheet = workbook.createSheet("Result Matrix");

            for (int i = 0; i < SIZE; i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < SIZE; j++) {
                    row.createCell(j).setCellValue(resultMatrix[i][j]);
                }
            }

            workbook.write(outputStream);
            System.out.println("Результаты сохранены в " + filePath);
        } catch (IOException e) {
            System.out.println("Ошибка при сохранении в Excel: " + e.getMessage());
        }
    }
}
