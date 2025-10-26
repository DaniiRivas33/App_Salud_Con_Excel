/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package ProyectoAppsalud;
/**
 *
 * @author danie
 */

import javax.swing.*;
import java.awt.*;
import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Salud extends JFrame {

    private JTextField txtPeso, txtAltura, txtIMC;
    private JButton btnCalcular, btnGuardar, btnLeer, btnInsertarFila;
    private JLabel lblTitulo;
    private final File archivoExcel = new File("src/main/java/ProyectoAppsalud/DatosSalud.xlsx");

    public Salud() {
        setTitle("Salud");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(500, 300);
        setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        //titulo
        
        lblTitulo = new JLabel("Salud - IMC", JLabel.CENTER);
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        add(lblTitulo, gbc);
        //campos peso, altura y IMC
        gbc.gridwidth = 1;
        gbc.gridy++;
        gbc.gridx = 0;
        add(new JLabel("Peso (kg):"), gbc);
        gbc.gridx = 1;
        txtPeso = new JTextField();
        add(txtPeso, gbc);

        gbc.gridy++;
        gbc.gridx = 0;
        add(new JLabel("Altura (m):"), gbc);
        gbc.gridx = 1;
        txtAltura = new JTextField();
        add(txtAltura, gbc);

        gbc.gridy++;
        gbc.gridx = 0;
        add(new JLabel("IMC:"), gbc);
        gbc.gridx = 1;
        txtIMC = new JTextField();
        txtIMC.setEditable(false);
        add(txtIMC, gbc);
        
        //Botones
        gbc.gridy++;
        gbc.gridx = 0;
        btnCalcular = new JButton("Calcular IMC");
        add(btnCalcular, gbc);
        
        gbc.gridx = 1;
        btnGuardar = new JButton("Guardar en Excel");
        add(btnGuardar, gbc);

        gbc.gridy++;
        gbc.gridx = 0;
        btnLeer = new JButton("Leer datos Excel");
        add(btnLeer, gbc);

        gbc.gridx = 1;
        btnInsertarFila = new JButton("Insertar fila");
        add(btnInsertarFila, gbc);

        // Acciones
        btnCalcular.addActionListener(e -> calcularIMC());
        btnGuardar.addActionListener(e -> guardarDatosEnExcel());
        btnLeer.addActionListener(e -> leerDatosExcel());
        btnInsertarFila.addActionListener(e -> insertarFila());

        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void calcularIMC() {
        try {
            double peso = Double.parseDouble(txtPeso.getText());
            double altura = Double.parseDouble(txtAltura.getText());
            double imc = peso / (altura * altura);
            txtIMC.setText(String.format("%.2f", imc));
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Introduce valores válidos (recuerda usar . en vez de, no añadas kg ni m)");
        }
    }

    private void guardarDatosEnExcel() {
    try (Workbook workbook = archivoExcel.exists()
            ? new XSSFWorkbook(new FileInputStream(archivoExcel))
            : new XSSFWorkbook()) {

        Sheet hoja = archivoExcel.exists()
                ? workbook.getSheetAt(0)
                : workbook.createSheet("Salud");

        if (hoja.getPhysicalNumberOfRows() == 0) {
            Row encabezado = hoja.createRow(0);
            encabezado.createCell(0).setCellValue("Peso (kg)");
            encabezado.createCell(1).setCellValue("Altura (m)");
            encabezado.createCell(2).setCellValue("IMC");    
        }

        int filaNueva = hoja.getLastRowNum() + 1;
        Row fila = hoja.createRow(filaNueva);
        fila.createCell(0).setCellValue(txtPeso.getText());
        fila.createCell(1).setCellValue(txtAltura.getText());
        fila.createCell(2).setCellValue(txtIMC.getText());

        
        hoja.autoSizeColumn(0);
        hoja.autoSizeColumn(1);
        hoja.autoSizeColumn(2);

        //imagen
        String imagePath = "src/main/java/Imagenes/images.jpg";
        File imgFile = new File(imagePath);
        if (imgFile.exists()) {
            try (InputStream inputStream = new FileInputStream(imagePath)) {
                byte[] bytes = inputStream.readAllBytes();

                int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

                Drawing<?> drawing = hoja.createDrawingPatriarch();
                CreationHelper helper = workbook.getCreationHelper();
                ClientAnchor anchor = helper.createClientAnchor();

                anchor.setCol1(3);
                anchor.setRow1(0);
                anchor.setCol2(5);
                anchor.setRow2(5);

                drawing.createPicture(anchor, pictureIdx);
            } catch (IOException ex) {
            }
        }  
        
        

        try (FileOutputStream out = new FileOutputStream(archivoExcel)) {
            workbook.write(out);
        }

        JOptionPane.showMessageDialog(this, "Datos guardados.");
    } catch (IOException ex) {
    }
}

    private void leerDatosExcel() {
        if (!archivoExcel.exists()) {
            JOptionPane.showMessageDialog(this, "El archivo no existe todavía");
            return;
        }

        try (FileInputStream fis = new FileInputStream(archivoExcel);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet hoja = workbook.getSheetAt(0);
            StringBuilder sb = new StringBuilder("Datos registrados:\n");

            for (Row fila : hoja) {
                for (Cell celda : fila) {
                    sb.append(String.format("%-25s", celda.toString())); //separa los numeros 25 caracteres para que no se junten todos
                }
                sb.append("\n");
            }

            JOptionPane.showMessageDialog(this, sb.toString());
        } catch (IOException ex) {
        }
    }

    private void insertarFila() {
        if (!archivoExcel.exists()) {
            JOptionPane.showMessageDialog(this, "Primero guarda algún dato en Excel.");
            return;
        }

        try (FileInputStream fis = new FileInputStream(archivoExcel);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet hoja = workbook.getSheetAt(0);
            int filaInsertar = 1; 
            hoja.shiftRows(filaInsertar, hoja.getLastRowNum(), 1);

            Row nuevaFila = hoja.createRow(filaInsertar);
            nuevaFila.createCell(0).setCellValue("Nueva Fila");
            nuevaFila.createCell(1).setCellValue("");
            nuevaFila.createCell(2).setCellValue("");

            try (FileOutputStream out = new FileOutputStream(archivoExcel)) {
                workbook.write(out);
            }

            JOptionPane.showMessageDialog(this, "Fila insertada correctamente.");
        } catch (IOException ex) {
        }
    }
    
    

    public static void main(String[] args) {
        SwingUtilities.invokeLater(Salud::new);
    }
}
