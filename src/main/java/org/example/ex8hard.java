package org.example;

import java.sql.*;
import java.io.*;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//абстрактный класс Student
abstract class Student {
    private String name;
    private int age;

    public Student(String name, int age) {
        this.name = name;
        this.age = age;
    }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public int getAge() { return age; }
    public void setAge(int age) { this.age = age; }

    public abstract void displayInfo();
}

//класс Worker (наследник Student)
class Worker extends Student {
    private double salary;

    public Worker(String name, int age, double salary) {
        super(name, age);
        this.salary = salary;
    }

    public double getSalary() { return salary; }
    public void setSalary(double salary) { this.salary = salary; }

    @Override //аннотация, переопределение метода суперкласса, наследование
    public void displayInfo() {
        System.out.println("Имя: " + getName() + ", Возраст: " + getAge() + ", Зарплата: " + salary);
    }
}

//работа с БД
class DatabaseManager {
    private static final String URL = "jdbc:mysql://localhost:3306/my_database";
    private static final String USER = "root";
    private static final String PASSWORD = "2dT#k9H!mQvL7pZ&bR1aWf";

    public static void createTable() {
        String sql = "CREATE TABLE IF NOT EXISTS workers (" +
                "id INT AUTO_INCREMENT PRIMARY KEY," +
                "name VARCHAR(50)," +
                "age INT," +
                "salary DOUBLE)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement()) {
            stmt.execute(sql);
            System.out.println("Таблица workers успешно создана.");
        } catch (SQLException e) {
            System.out.println("Ошибка создания таблицы: " + e.getMessage());
        }
    }

    public static void saveWorkerToDB(Worker worker) {
        String sql = "INSERT INTO workers (name, age, salary) VALUES (?, ?, ?)";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, worker.getName());
            pstmt.setInt(2, worker.getAge());
            pstmt.setDouble(3, worker.getSalary());
            pstmt.executeUpdate();
            System.out.println("Работник успешно сохранен в БД.");
        } catch (SQLException e) {
            System.out.println("Ошибка сохранения данных: " + e.getMessage());
        }
    }

    public static void displayWorkers() {
        String sql = "SELECT * FROM workers";
        try (Connection conn = DriverManager.getConnection(URL, USER, PASSWORD);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                System.out.println("ID: " + rs.getInt("id") +
                        ", Имя: " + rs.getString("name") +
                        ", Возраст: " + rs.getInt("age") +
                        ", Зарплата: " + rs.getDouble("salary"));
            }
        } catch (SQLException e) {
            System.out.println("Ошибка при выводе данных: " + e.getMessage());
        }
    }
}

//работа с Excel
class ExcelManager {
    private static final String FILE_PATH = "workers.xlsx";

    public static void saveToExcel(Worker worker) {
        Workbook workbook;
        Sheet sheet;
        File file = new File(FILE_PATH);

        try {
            if (file.exists()) {
                try (FileInputStream fis = new FileInputStream(file)) {
                    workbook = new XSSFWorkbook(fis);
                }
            } else {
                workbook = new XSSFWorkbook();
            }

            sheet = workbook.getSheet("Работники");
            if (sheet == null) {
                sheet = workbook.createSheet("Работники");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Имя");
                header.createCell(1).setCellValue("Возраст");
                header.createCell(2).setCellValue("Зарплата");
            }

            int rowNum = sheet.getLastRowNum() + 1;
            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue(worker.getName());
            row.createCell(1).setCellValue(worker.getAge());
            row.createCell(2).setCellValue(worker.getSalary());

            try (FileOutputStream fos = new FileOutputStream(FILE_PATH)) {
                workbook.write(fos);
            }
            System.out.println("Данные сохранены в Excel.");
        } catch (IOException e) {
            System.out.println("Ошибка записи в Excel: " + e.getMessage());
        }
    }
}

//главный класс
public class ex8hard {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        DatabaseManager.createTable();

        while (true) {
            System.out.println("\nВыберите действие:");
            System.out.println("1. Добавить работника");
            System.out.println("2. Вывести всех работников");
            System.out.println("3. Сохранить в Excel");
            System.out.println("0. Выход");
            System.out.print("Ваш выбор: ");

            int choice;
            try {
                choice = scanner.nextInt();
            } catch (Exception e) {
                System.out.println("Ошибка ввода! Введите число.");
                scanner.nextLine();
                continue;
            }

            switch (choice) {
                case 1:
                    scanner.nextLine();
                    System.out.print("Введите имя: ");
                    String name = scanner.nextLine();

                    System.out.print("Введите возраст: ");
                    int age;
                    while (true) {
                        if (scanner.hasNextInt()) {
                            age = scanner.nextInt();
                            break;
                        } else {
                            System.out.println("Ошибка! Введите целое число.");
                            scanner.next();
                        }
                    }

                    System.out.print("Введите зарплату: ");
                    double salary;
                    while (true) {
                        if (scanner.hasNextDouble()) {
                            salary = scanner.nextDouble();
                            break;
                        } else {
                            System.out.println("Ошибка! Введите число.");
                            scanner.next();
                        }
                    }

                    Worker worker = new Worker(name, age, salary);
                    DatabaseManager.saveWorkerToDB(worker);
                    ExcelManager.saveToExcel(worker);
                    break;

                case 2:
                    DatabaseManager.displayWorkers();
                    break;

                case 3:
                    System.out.println("Все данные уже сохранены в Excel.");
                    break;

                case 0:
                    System.out.println("Выход из программы.");
                    scanner.close();
                    return;

                default:
                    System.out.println("Некорректный ввод. Введите число от 0 до 3");
            }
        }
    }
}
