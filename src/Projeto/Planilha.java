package Projeto;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Planilha {
	public static void main(String[] args ) throws Throwable {
		
		File file = new File("C:\\Users\\angel\\eclipse-workspace\\Curso_Jdev_Avancado2\\src\\Projeto\\Planilha.xlsx");
		
		if (!file.exists()) {
			file.createNewFile();
			
		}
		
		/*SETANDO DADOS*/
		
		Pessoa pessoas = new Pessoa();
		pessoas.setNome("ANGELO");
		pessoas.setPais("BRASIL");
		pessoas.setIdade(25);
		Pessoa pessoas1 = new Pessoa();
		pessoas1.setNome("ANDRE");
		pessoas1.setPais("ESPANHA");
		pessoas1.setIdade(29);
		
		/* CRIANDO LISTA*/
		
		List<Pessoa> Lpessoas = new ArrayList<>();
		Lpessoas.add(pessoas);
		Lpessoas.add(pessoas1);
		

		/* CRIAR UMA NOVA PLANILHA*/
		Workbook worlbook = new XSSFWorkbook();/*classe da biblioteca Apache POI usada para manipular arquivos Excel no formato .xlsx (Excel 2007 e versões posteriores*/
		Sheet LinhaPlanilha = (Sheet) worlbook.createSheet("planilha criada"); //RESPONSAVEL POR CIRAR A PLANILHA
		
		
		int linhas = 0;
		
		for(Pessoa p: Lpessoas){
			
			Row linha =  LinhaPlanilha.createRow(linhas ++); /* CRIAÇÃO DE LINHA NA PLANILHA*/
			
			int celula = 0;
			
			Cell celNome = linha.createCell(celula ++);
			celNome.setCellValue(p.getNome());
			
			Cell celPais = linha.createCell(celula ++);
			celPais.setCellValue(p.getPais());
			
			Cell celIdade = linha.createCell(celula ++);
			celIdade.setCellValue(p.getIdade());
		
		}/* FIM DA CRIAÇÃO DA PLANILHA*/
		
			/*SALVAR A PLANILHA EM UM ARQUIVO*/
			try {
				FileOutputStream arquivo = new FileOutputStream(file);
				worlbook.write(arquivo);
				arquivo.flush();
				arquivo.close();
			} catch (FileNotFoundException e) {
				
				e.printStackTrace();
				System.out.println("erro ao criar planilha");
			}
				

        System.out.println("Planilha criada com sucesso!");
    }

}
