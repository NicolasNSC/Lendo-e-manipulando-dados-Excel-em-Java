import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;


public class LerExcel {

	private static final String fileName = ""; //caminho do diretorio do arquivo

	public static void main(String[] args) throws IOException {

		List<Alunos> listaAlunos = new ArrayList<Alunos>();

		try {
			FileInputStream arquivo = new FileInputStream(new File(LerExcel.fileName));

			XSSFWorkbook workbook = new XSSFWorkbook(arquivo);

			XSSFSheet sheetAlunos = workbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheetAlunos.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				Alunos aluno = new Alunos();
				listaAlunos.add(aluno);
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getColumnIndex()) {
						case 0:
							aluno.setNome(cell.getStringCellValue());
							break;
						case 1:
							aluno.setRa(String.valueOf(cell.getNumericCellValue()));
							break;
						case 2:
							aluno.setNota1(cell.getNumericCellValue());
							break;
						case 3:
							aluno.setNota2(cell.getNumericCellValue());
							break;
						case 4:
							aluno.setMedia(cell.getNumericCellValue());
							break;
						case 5:
							aluno.setAprovado(cell.getBooleanCellValue());
							break;
					}
				}
			}
			arquivo.close();
			workbook.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		}

		if (listaAlunos.size() == 0) {
			System.out.println("Nenhum aluno encontrado!");
		} else {
			double soma = 0;
			double maior = 0;
			double menor = listaAlunos.get(0).getMedia();
			int aprovados = 0;
			int reprovados = 0;
			int mediana = 8;

			System.out.println("Notas semestrais\n");

			for (Alunos aluno : listaAlunos) {
				System.out.println(aluno.getNome() + " -  Média: " + aluno.getMedia());

				soma = soma + aluno.getMedia();
				if (aluno.getMedia() > maior) {
					maior = aluno.getMedia();
				}
				if (aluno.getMedia() < menor) {
					menor = aluno.getMedia();
				}
				if (aluno.getMedia() >= 6) {
					aprovados++;
				}
				if (aluno.getMedia() < 6) {
					reprovados++;
				}
			}

			Map<Double, Integer> frequenciaNumeros = new HashMap<>();

			int maiorFrequencia = 0;

			for (Alunos aluno : listaAlunos) {


				Integer quantidade = frequenciaNumeros.get(aluno.getMedia());
				if (quantidade == null) {
					quantidade = 1;
				} else {
					quantidade += 1;
				}
				frequenciaNumeros.put(aluno.getMedia(), quantidade);

				if (maiorFrequencia < quantidade) {
					maiorFrequencia = quantidade;
				}
			}
			System.out.print("\nA(s) moda(s) é (são): ");
			for (Double numeroChave : frequenciaNumeros.keySet()) {
				int quantidade = frequenciaNumeros.get(numeroChave);
				if (maiorFrequencia == quantidade) {
					System.out.print(numeroChave + " ");
				}
			}
			double media = soma / listaAlunos.size();

			System.out.printf("\nA media de notas é: %.2f", media);
			System.out.println("\nA mediana é: " + mediana);
			System.out.println("A maior nota é: " + maior);
			System.out.println("A menor nota é: " + menor);
			System.out.println("O numero de alunos aprovados é: " + aprovados);
			System.out.println("O numero de alunos reprovados é: " + reprovados);
			System.out.println("Número total de alunos: " + listaAlunos.size());


		}


	}
}

