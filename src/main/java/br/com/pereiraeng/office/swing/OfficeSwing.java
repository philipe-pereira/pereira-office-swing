package br.com.pereiraeng.office.swing;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.swing.AbstractListModel;
import javax.swing.JTable;
import javax.swing.table.AbstractTableModel;
import javax.swing.tree.DefaultMutableTreeNode;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.pereiraeng.math.swing.chart.Chart;
import br.com.pereiraeng.math.swing.chart.Cloud;
import br.com.pereiraeng.math.swing.chart.CurveFamily;
import br.com.pereiraeng.math.swing.chart.Plotable;
import br.com.pereiraeng.office.Office;
import br.com.pereiraeng.swing.table.treetable.model.AdvancedTreeTableModel;

public class OfficeSwing {

	// ---------------------- XSSF - TableModel ----------------------

	/**
	 * Função que transforma uma tabela gráfica {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file       arquivo a ser criado com a planilha Excel
	 * @param tableModel modelo de tabela contendo o conteúdo e o cabeçalho das
	 *                   colunas
	 */
	public static void export(File file, AbstractTableModel tableModel) {
		export(file, "001", tableModel);
	}

	public static void export(File file, String sheetName, AbstractTableModel tableModel) {
		export(file, sheetName, tableModel, null);
	}

	/**
	 * Função que transforma uma tabela gráfica {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file           arquivo a ser criado com a planilha Excel
	 * @param sheetName      nome da folha da planilha
	 * @param tableModel     modelo de tabela contendo o conteúdo da tabela e o
	 *                       cabeçalho das colunas
	 * @param rowHeaderModel modelo da lista contendo o cabeçalho das linhas
	 */
	public static void export(File file, String sheetName, AbstractTableModel tableModel,
			AbstractListModel<?> rowHeaderModel) {
		export(file, new String[] { sheetName }, new AbstractTableModel[] { tableModel },
				new AbstractListModel[] { rowHeaderModel });
	}

	/**
	 * Função que transforma tabelas gráficas {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file            arquivo a ser criado com a planilha Excel
	 * @param tableModels     vetor de modelos de tabelas contendo o conteúdo e o
	 *                        cabeçalho das colunas
	 * @param rowHeaderModels vetor de modelos da lista contendo o cabeçalho das
	 *                        linhas
	 */
	public static void export(File file, String[] sheetNames, AbstractTableModel[] tableModels,
			AbstractListModel<?>[] rowHeaderModels) {
		XSSFWorkbook wb = Office.getWB(file);

		for (int k = 0; k < tableModels.length; k++)
			writeSheet(wb, sheetNames[k], tableModels[k], rowHeaderModels[k]);

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static XSSFSheet writeSheet(XSSFWorkbook wb, String sheetName, AbstractTableModel tableModel,
			AbstractListModel<?> rowHeaderModel) {
		// se houver o cabeçalho das linhas, deslocar uma coluna para ele
		int rh = rowHeaderModel != null ? 1 : 0;

		XSSFSheet sh = wb.createSheet(sheetName);

		// cabeçalho das colunas
		XSSFRow row = sh.createRow(0);
		for (int i = 0; i < tableModel.getColumnCount(); i++) {
			XSSFCell cell = row.createCell(i + rh);
			String columnName = tableModel.getColumnName(i);
			cell.setCellValue(columnName);
		}

		for (int j = 0; j < tableModel.getRowCount(); j++) {
			row = sh.createRow(j + 1);

			// cabeçalho das linhas
			if (rowHeaderModel != null)
				Office.setCell(row.createCell(0), rowHeaderModel.getElementAt(j));

			// conteúdo
			for (int i = 0; i < tableModel.getColumnCount(); i++)
				Office.setCell(row.createCell(i + rh), tableModel.getValueAt(j, i));
		}

		return sh;
	}

	// =========================== TREE-TABLE ===========================

	/**
	 * Função que exporta uma tabela-árvore para uma planilha Excel (a arborescência
	 * será convertida em uma série de linhas, onde os nós superiores mesclam seus
	 * filhos)
	 * 
	 * @param file       arquivo a ser criado com a planilha do MS-Excel
	 * @param sheetName  nome da folha da planilha
	 * @param attm       modelo de tabela-árvore contendo o conteúdo da tabela e o
	 *                   cabeçalho das colunas
	 * @param treeLevels nome das colunas da árvore (o tamanho deste vetor deve ser
	 *                   igual à profundidade da árvore)
	 */
	public static void export(File file, String sheetName, AdvancedTreeTableModel attm, String... treeLevels) {
		String[] valueColumns = new String[attm.getColumnCount() - 1];
		for (int j = 0; j < valueColumns.length; j++)
			valueColumns[j] = attm.getColumnName(j + 1);
		export(file, sheetName, (DefaultMutableTreeNode) attm.getRoot(), attm.getTableData(), valueColumns, treeLevels);
	}

	/**
	 * Função que exporta uma tabela-árvore para uma planilha Excel (a arborescência
	 * será convertida em uma série de linhas, onde os nós superiores mesclam seus
	 * filhos)
	 * 
	 * @param file         arquivo a ser criado com a planilha Excel do MS-Excel
	 * @param sheetName    nome da folha da planilha
	 * @param treeLevels   nome das colunas da árvore (o tamanho deste vetor deve
	 *                     ser igual à profundidade da árvore)
	 * @param root         raiz da árvore
	 * @param data         tabela com os dados de cada nó
	 * @param valueColumns demais colunas, respectivas aos valores de cada nó (pode
	 *                     ser <code>null</code>, e neste caso não haverá cabeçalho
	 *                     para os valores)
	 */
	public static void export(File file, String sheetName, DefaultMutableTreeNode root, Map<Object, Object[]> data,
			String[] valueColumns, String... treeLevels) {
		int depth = root.getDepth();

		// workbook e sheet
		XSSFWorkbook wb = Office.getWB(file);

		XSSFSheet sh = wb.createSheet(sheetName);
		sh.createFreezePane(depth, 1);

		// cabeçalho da coluna da árvore
		XSSFRow row = sh.createRow(0);

		XSSFCell cell = null;
		for (int k = 0; k < treeLevels.length; k++) {
			cell = row.createCell(k);
			cell.setCellValue(treeLevels[k]);
		}

		if (valueColumns != null)
			for (int j = 0; j < valueColumns.length; j++) {
				cell = row.createCell(treeLevels.length + j);
				cell.setCellValue(valueColumns[j]);
			}

		// conteúdo
		DefaultMutableTreeNode[] nodes = new DefaultMutableTreeNode[depth];
		int[] rst = new int[depth];
		int[] starts = new int[depth];
		writeLine(sh, root, data, 1, nodes, rst, starts);

		// escrever arquivo
		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException exc) {
			exc.printStackTrace();
		}
	}

	private static int writeLine(XSSFSheet sh, DefaultMutableTreeNode node, Map<Object, Object[]> data, int rowIndex,
			DefaultMutableTreeNode[] nodes, int[] rst, int[] starts) {
		if (node.isLeaf()) {
			// se for uma folha, contém valores...

			int level = node.getLevel();
			Object obj = node.getUserObject();

			Object[] values = data.get(obj);

			XSSFRow row = sh.createRow(rowIndex);

			// folha
			XSSFCell cell = row.createCell(level - 1);
			Office.setCell(cell, obj);

			// medições
			for (int m = 0; m < values.length; m++) {
				cell = row.createCell(level + m);
				Office.setCell(cell, values[m]);
			}

			// valor da célula mesclada
			for (int i = rst.length - 1; i > 0; i--) {
				if (rst[i] == 0) { // quando a coluna atual começa...
					starts[i - 1] = rowIndex;
					cell = row.createCell(i - 1);
					cell.setCellValue(nodes[i - 1].toString());
				} else
					break;
			}

			// mesclagem
			for (int i = rst.length - 1; i > 0; i--) {
				if (rst[i] == nodes[i - 1].getChildCount() - 1) {
					// quando a coluna atual termina...
					if (starts[i - 1] != rowIndex)
						sh.addMergedRegion(new CellRangeAddress(starts[i - 1], rowIndex, i - 1, i - 1));
				} else
					break;
			}
			rowIndex++;
			return rowIndex;
		} else {
			// se o nó possuir filhos...
			for (int l = 0; l < node.getChildCount(); l++) {
				// ler recursivamente os filhos
				DefaultMutableTreeNode child = (DefaultMutableTreeNode) node.getChildAt(l);
				int level = child.getLevel() - 1;
				nodes[level] = child;
				rst[level] = l;
				rowIndex = writeLine(sh, child, data, rowIndex, nodes, rst, starts);
			}
			return rowIndex;
		}
	}

	// ============================= CHART -> XLSX =============================

	public static void export(File file, Chart<?> chart) {
		XSSFWorkbook wb = new XSSFWorkbook();

		// nomes das etiquetas
		List<?> labels = chart.getKeyArray();

		for (int l = 0; l < labels.size(); l++) {
			XSSFSheet sh = wb.createSheet(labels.get(l).toString());

			XSSFRow row = sh.createRow(0);
			row.createCell(0).setCellValue("X");
			row.createCell(1).setCellValue("Y");

			Object obj = (Object) labels.get(l);
			Plotable plotable = chart.get(obj);
			if (plotable instanceof Cloud) {
				Cloud c = (Cloud) plotable;
				double[][] xy = c.getCoordinates();

				for (int j = 0; j < xy[0].length; j++) {
					row = sh.createRow(j + 1);
					row.createCell(0).setCellValue(xy[0][j]);
					row.createCell(1).setCellValue(xy[1][j]);
				}
			} else if (plotable instanceof CurveFamily) {
				CurveFamily cf = (CurveFamily) plotable;
				for (int k = 0; k < cf.size(); k++) {
					cf.setIndex(k);
					double[][] xy = cf.getCoordinates();

					for (int j = 0; j < xy[0].length; j++) {
						row = sh.createRow(j + 1);
						row.createCell(2 * k).setCellValue(xy[0][j]);
						row.createCell(2 * k + 1).setCellValue(xy[1][j]);
					}
				}
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
