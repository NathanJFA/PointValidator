package com.code.point_validator;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComparatorSheets {

    public ComparatorSheets(){};

    public void sysMain(String pathPlanilhaEntrada, String pathPlanilhaValidacao) {

        if(pathPlanilhaEntrada.isBlank() || pathPlanilhaValidacao.isBlank()) {
            return;
        }
        // Abrir as planilhas de entrada e saída
        FileInputStream fileEntrada = null;
        FileInputStream fileValidacao = null;

        XSSFWorkbook workbookPlanilhaPontoEletronico = null;
        XSSFWorkbook workbookPlanilhaValidacao = null;
        try {
            fileEntrada = new FileInputStream(pathPlanilhaEntrada);
            fileValidacao = new FileInputStream(pathPlanilhaValidacao);
            workbookPlanilhaPontoEletronico = new XSSFWorkbook(fileEntrada);
            workbookPlanilhaValidacao = new XSSFWorkbook(fileValidacao);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        XSSFSheet sheetEntrada = workbookPlanilhaPontoEletronico.getSheetAt(0);
        XSSFSheet sheetComparacao = workbookPlanilhaValidacao.getSheetAt(0);

        //AJUSTAR DAQUI PRA BAIXO
        // Identificar a coluna da data de entrada
        int colunaDataEntrada = 2; // exemplo: data de entrada na coluna C

        // Percorrer as linhas das duas planilhas ao mesmo tempo
        for (int i = 0; i <= sheetEntrada.getLastRowNum() && i <= sheetComparacao.getLastRowNum(); i++) {
            XSSFRow rowEntrada1 = sheetEntrada.getRow(i);
            XSSFRow rowEntrada2 = sheetComparacao.getRow(i);

            if (rowEntrada1 != null && rowEntrada2 != null) {
                // Comparar as informações do funcionário
                XSSFCell cellEntrada = rowEntrada1.getCell(colunaDataEntrada);
                XSSFCell cellSaida = rowEntrada2.getCell(colunaDataEntrada);

                if (cellEntrada != null && cellSaida != null && isMesmoDia(cellEntrada.getDateCellValue(), new Date())) {
                    // Marcar a presença do funcionário na planilha de saída
                    Cell cellSaidaPresenca = rowEntrada2.getCell(colunaDataEntrada + 1);
                    if (cellSaidaPresenca == null) {
                        cellSaidaPresenca = rowEntrada2.createCell(colunaDataEntrada + 1);
                    }
                    cellSaidaPresenca.setCellValue("Presente");
                }
            }
        }

        // Salvar a planilha modificada
        FileOutputStream arquivoSaidaModificado = null;
        try {
            arquivoSaidaModificado = new FileOutputStream("planilha_saida_modificada.xlsx");
            workbookPlanilhaValidacao.write(arquivoSaidaModificado);
            arquivoSaidaModificado.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


        // Fechar os arquivos e liberar recursos
        try {
            workbookPlanilhaPontoEletronico.close();
            workbookPlanilhaValidacao.close();
            fileEntrada.close();
            fileValidacao.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

        // Verifica se duas datas são do mesmo dia
    private static boolean isMesmoDia(Date data1, Date data2) {
        Calendar cal1 = Calendar.getInstance();
        cal1.setTime(data1);
        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(data2);
        return cal1.get(Calendar.YEAR) == cal2.get(Calendar.YEAR) &&
                cal1.get(Calendar.MONTH) == cal2.get(Calendar.MONTH) &&
                cal1.get(Calendar.DAY_OF_MONTH) == cal2.get(Calendar.DAY_OF_MONTH);
    }
}
