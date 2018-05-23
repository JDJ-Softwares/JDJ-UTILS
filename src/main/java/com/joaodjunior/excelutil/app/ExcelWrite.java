package com.joaodjunior.excelutil.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	
	private String filePath;
	private String fileName;
	
	private List<String> sheetsName;
	
	private Object object;
	
	private XSSFWorkbook workbook;
	
	private XSSFSheet sheet;

	public ExcelWrite() {
		super();
	}
	
	public static void main(String[] args) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, InstantiationException {
		ExcelWrite w = new ExcelWrite();
		Pessoa teste = new Pessoa("Joao", "Alo");
		List<Object> ob = new ArrayList<Object>();
		try {
			ob = w.readExcel(new File(System.getProperty("user.home").concat("/Downloads/Teste.xls")), Pessoa.class, Endereco.class);
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		for(Object pes : ob) {
			if(pes.getClass() == Pessoa.class) {
				Pessoa p = (Pessoa) pes;
				System.out.println(p.getNome() +" "+p.getSobrenome());				
			} else {
				Endereco p = (Endereco) pes;
				System.out.println(p.getRua() +" "+p.getNumero());
			}
		}
		
		w.createExcelFile(ob, Pessoa.class, Endereco.class);
		
	}
	
	public <T> void createExcelFile(List<Object> objects, Class<?>... model) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, InstantiationException {
		
		int rowNum = 0;
		Map<String, Class<?>> map = new HashMap<String, Class<?>>();
		for (Class<?> c : model) {
			map.put(c.getName(), c);
		}
		Row row = getSheet().createRow(rowNum++);
		int colNum = 0;
		for(Object object : objects) {			
			Class<?> clazz = map.get(object.getClass().getName());
			
			//Cria o Cabecalho com base no modelo
			if(rowNum == 1 ) {
				for(Class<?> cl : model) {
					for(Field field : cl.getDeclaredFields()) {	//Colocar esse FOR fora para criar cabecalho
						Cell cell = row.createCell(colNum++);
						if (field.getType().getSimpleName().equals("String")) {
							cell.setCellValue((String) field.getName().toUpperCase());
						} else {
							cell.setCellValue((String) field.getName().toUpperCase());
						}
					}					
				}
				row = getSheet().createRow(rowNum++);
				colNum = 0;
			}
			if(clazz != null){
				for(Field field : clazz.getDeclaredFields()) {
					Cell cell = row.createCell(colNum++);
					
					Method m = buscarGetter(field.getName(), clazz.getDeclaredMethods());
					
					if(field.getType() == String.class) {
						cell.setCellValue((String) m.invoke(object, null) );
					} else if (field.getType() == int.class) {
						cell.setCellValue((String) String.valueOf(object));
					}
				}				
			}
			
			if(object.getClass() == objects.get(objects.size() - 1).getClass()) {
				row = getSheet().createRow(rowNum++);
				colNum = 0;
			}
			
		}
		
		FileOutputStream outputStream = null;
		
		try {
			if(!getFileName().isEmpty()) {
				outputStream = new FileOutputStream(verifyAndCreateFileData());
			} else {
				outputStream = new FileOutputStream(getFileName());				
			}
            getWorkbook().write(outputStream);
            getWorkbook().close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }
	
	private Method buscarGetter(String name, Method[] methods) {
		
		Method method = null;
		
		for(Method m : methods) {
			String methodNome = null;
			String nome = m.getName().trim().substring(3).toLowerCase();
			String funcao = m.getName().trim().substring(0, 3);
			if(nome.equals(name) && funcao.equals("get")) {
				method = m;
				break;
			}
		}
		
		return method;
	}

	public <T> List<Object> readExcel(File file, Class<?>... model) throws IOException, NoSuchMethodException, SecurityException {
		
		List<Object> objects = new ArrayList<Object>();
		
		setFilePath(file.getAbsolutePath());
		setFileName(file.getName());
		
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
		Sheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet.rowIterator();
		
		while(rowIterator.hasNext()) {
			Row row = (Row) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			//Ler o cabecalho das colunas
			
			for(Class<?> m : model) {
				int i = 0;
				List<Field> fieldsClass = readClass(m);
				Map<String, Object> map = new HashMap<String, Object>();
				while(i < fieldsClass.size() && cellIterator.hasNext()) {
					Cell cell = (Cell) cellIterator.next();
					if ((cell.getCellTypeEnum() == CellType.STRING) && (fieldsClass.get(i).getType() == String.class)) {
						map.put(fieldsClass.get(i).getName(), cell.getStringCellValue());					
					} else if((cell.getCellTypeEnum() == CellType.NUMERIC) && (fieldsClass.get(i).getType() == int.class)) {
						map.put(fieldsClass.get(i).getName(), ((Double)cell.getNumericCellValue()).intValue());
					} else if((cell.getCellTypeEnum() == CellType.BOOLEAN) && (fieldsClass.get(i).getType() == boolean.class)) {
						map.put(fieldsClass.get(i).getName(), cell.getBooleanCellValue());
					} else if((cell.getCellTypeEnum() == CellType.FORMULA) || (fieldsClass.get(i).getType() == Date.class)) {
						map.put(fieldsClass.get(i).getName(), cell.getDateCellValue());
					} else {
						continue;
					}
					i++;
				}
				try {
					if(!map.isEmpty()) {
						objects.add(m.getConstructor(Map.class).newInstance(map));						
					}
				} catch (InstantiationException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				}
			}
		}
		workbook.close();
		
		return objects;
		
	}
	
	private <T> List<Field> readClass(Class<T> model) {
		
		List<Field> fieldsClass = new ArrayList<Field>();
		for(Field field : model.getDeclaredFields()) {
			fieldsClass.add(field);
		}
		return fieldsClass;
	}
		
	private File verifyAndCreateFileData() {
		int i = 0;
		File file = new File(getFilePath()+getFileName());
		
		while(file.exists()) {
			file = new File(getFilePath()+getFileName()+String.valueOf(i));
			i++;
		}
		
		return file;

	}
	
	private static Set<String> getMethodNames(Method[] methods) {
        Set<String> membersNames = new TreeSet<String>();
        for (Method member: methods) {
            membersNames.add(member.getName());
        }
        return  membersNames;
    }
	
	public XSSFWorkbook getWorkbook() {
		if(workbook == null) {
			setWorkbook(new XSSFWorkbook());
			return workbook;
		} else {
			return workbook;			
		}
	}

	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public XSSFSheet getSheet() {
		if(sheet == null) {
			sheet = getWorkbook().createSheet();
			return sheet;
		} else {
			return sheet;			
		}
	}

	public void setSheet(XSSFSheet sheet) {
		this.sheet = sheet;
	}

	public String getFilePath() {
		if(filePath == null) {
			setFilePath(System.getProperty("user.home").concat("/Downloads/"));
		}
		return filePath;
	}

	public Object getObject() {
		return object;
	}

	public void setObject(Object object) {
		this.object = object;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public String getFileName() {
		if(fileName == null) {
			setFileName("GenericFile.xlsx");
			return fileName;
		} else {
			return fileName;			
		}
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public List<String> getSheetsName() {
		return sheetsName;
	}

	public void setSheetsName(List<String> sheetsName) {
		this.sheetsName = sheetsName;
	}
}