package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class InputMahasiswa {
    // Menyimpan nama sesi ini untuk validasi duplikat
    private static ArrayList<String> namaMahasiswa = new ArrayList<>();
    // Path file sesuai screenshot
    private static final String FILE_PATH = "src/main/resources/data_mahasiswa.xlsx";

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.println("Masukkan data mahasiswa. Ketik 'selesai' pada nama untuk mengakhiri");

        while (true) {
            System.out.print("Masukkan Nama: ");
            String nama = scanner.nextLine();

            // Cek kondisi keluar
            if (nama.equalsIgnoreCase("selesai")) {
                break;
            }

            // Validasi: Cek apakah nama sudah ada
            if (namaMahasiswa.contains(nama)) {
                System.out.println("Nama sudah ada, masukkan nama yang berbeda !");
                continue; // Kembali ke awal loop
            }

            System.out.print("Masukkan Semester: ");
            int semester = 0;
            try {
                semester = Integer.parseInt(scanner.nextLine());
            } catch (NumberFormatException e) {
                System.out.println("Semester harus angka!");
                continue;
            }

            System.out.print("Masukkan Mata Kuliah: ");
            String mataKuliah = scanner.nextLine();

            // Simpan data ke list (untuk validasi nama selanjutnya)
            namaMahasiswa.add(nama);

            // Simpan data ke Excel
            simpanKeExcel(nama, semester, mataKuliah);
        }

        System.out.println("Terima kasih !");
    }

    private static void simpanKeExcel(String nama, int semester, String mataKuliah) {
        Workbook workbook;
        Sheet sheet;

        File file = new File(FILE_PATH);

        try {
            // Cek apakah file sudah ada
            if (file.exists()) {
                // Jika ada, buka file tersebut (Mode Append)
                FileInputStream fis = new FileInputStream(file);
                workbook = new XSSFWorkbook(fis);
                sheet = workbook.getSheetAt(0);
                fis.close();
            } else {
                // Jika tidak ada, buat file baru
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Data Mahasiswa");

                // Buat Header (Baris 0)
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Nama");
                headerRow.createCell(1).setCellValue("Semester");
                headerRow.createCell(2).setCellValue("Mata Kuliah");
            }

            // Menentukan baris terakhir yang kosong
            int lastRowNum = sheet.getLastRowNum();
            Row row = sheet.createRow(lastRowNum + 1);

            // Isi data ke kolom
            row.createCell(0).setCellValue(nama);
            row.createCell(1).setCellValue(semester);
            row.createCell(2).setCellValue(mataKuliah);

            // Tulis perubahan ke file
            FileOutputStream fos = new FileOutputStream(FILE_PATH);
            workbook.write(fos);
            workbook.close();
            fos.close();

            System.out.println("Data berhasil disimpan ke dalam file data_mahasiswa.xlsx !");
            System.out.println(); // Baris baru agar rapi

        } catch (IOException e) {
            System.out.println("Terjadi kesalahan saat menyimpan file: " + e.getMessage());
        }
    }
}