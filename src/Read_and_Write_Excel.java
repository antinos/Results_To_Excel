//import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
///import java.lang.reflect.Array;				//v1.0.2
///import java.util.ArrayList;					//v1.0.2
///import java.util.Arrays;						//v1.0.2
///import java.nio.channels.FileChannel;		//If creating workbook from file, use to create ReadFile copy
import java.io.InputStream;						//Required if reading the excel file as a Fileinputstream
///import java.util.StringJoiner;					//for data in string
import java.io.OutputStream;

//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellStyle;		//Try without setting style
import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.Font;			//Try without setting style
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.WorkbookFactory;	//oringally used to create workbook before xlsx was needed for later sxssf export.
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;		//streamable xlsx workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook;		//non-streamable xlsx workbook

import ij.IJ;
import ij.Macro;
import ij.measure.ResultsTable;
import ij.plugin.PlugIn;
import ij.plugin.filter.Analyzer;

// below is public class header for troubleshooting (with it, you can run the program in the IDE).
///public class Read_and_Write_Excel {
///public static void main(String[] args) {

public class Read_and_Write_Excel implements PlugIn {
public void run(String arg) {
	
	String image_title = null;
	String Str1 = Macro.getOptions();
	int counter = 0;							//for use with 1D arrays instead of 2D

	if (Str1 != null) {
	image_title = Str1;

	}
	
	//Replace spaces with '_' underscores in the results table
	String[] ColHeadings = Analyzer.getResultsTable().getHeadings();
	
	String RA = new String(); 
	///Clipboard SysClip = Toolkit.getDefaultToolkit().getSystemClipboard();	//can be used for macro to communicate with plugin and visa versa
	int ColHedLen = ColHeadings.length;
	int RTcount = Analyzer.getResultsTable().getCounter();
	///String data_in_matrix[][] = new String[ColHedLen][RTcount];	//for data in 2D matrix
	///StringJoiner data_in_string = new StringJoiner(" ");			//for data in string
	String data_in_1Darray[] = new String[ColHedLen*RTcount];		//for data in 1D array
	///ArrayList<String> data_in_list = new ArrayList<String>();	//for data in ArrayList
	ResultsTable rt = Analyzer.getResultsTable();
	for (int y=0; y<RTcount; y++) {
		for (int x=0; x<ColHeadings.length; x++) {
			RA = String.valueOf(rt.getStringValue(ColHeadings[x], y));
			
			//Below 8 lines... alternate method to get RT data via a macro and the system clipboard
			///IJ.runMacro("var TempRT = getResultString("+ (char)34 + ColHeadings[x] + (char)34 + ", " + y + ");\nString.copy(TempRT);");
			///try {
				///RA = (String) SysClip.getData(DataFlavor.stringFlavor);
			///} catch (UnsupportedFlavorException e) {
				///e.printStackTrace();
			///} catch (IOException e) {
				///e.printStackTrace();
			///}
			
			RA = RA.replaceAll("\\s+", "_");
			///data_in_matrix[x][y] = RA;					//data in 2D matrix
			///data_in_string.add(RA);						//data in string
			data_in_1Darray[counter] = RA;					//data in 1D Array
				counter++;
			///data_in_list.add(RA);						//data in ArrayList
			}
		}
	
/*	//block-comment below = working code from v1.0.2

	//below... alternate method to extract data from the results table
  	int RowN_alt = Analyzer.getResultsTable().getCounter();
  	StringBuilder ResultsSB = new StringBuilder();
  	for(int i=0; i<RowN_alt; i++) {
  		String line = Analyzer.getResultsTable().getRowAsString(i);
  		//when you getRowAsString it seems to also extract count values. Below 7 lines is code to remove it. 
  		String[] line_array = line.trim().split("\\s+");
  		ArrayList<String> array_list = new ArrayList<String>(Arrays.asList(line_array));
  		//above... need to pass the array into a new ArrayList otherwise below... remove would not work.
  		array_list.remove(0);
  		String line_list = "";
  		for (String s : array_list){
  			line_list += s + "\t";
  		} //above... and finally back to a string again. Maybe some other way to integrate StringBuilder but did not try to find it.
    	ResultsSB.append(line_list + " ");
    }
  	String Results_as_string = ResultsSB.toString();
    String[] Results_in_array = Results_as_string.trim().split("\\s+");
    ////data_in_array = new Double[Results_in_array.length];
    ////for(int i = 0; i < Results_in_array.length; i++)
    ////{
    	////data_in_array[i] = Double.parseDouble(Results_in_array[i]);
    ////}
	
	int ArrayL = Array.getLength(Results_in_array);
	String[] ColHead_in_array = Analyzer.getResultsTable().getHeadings();
	//final ResultsTable rt = Analyzer.getResultsTable(); this code didn't work. Replaced with above.
	//String ColHead_string  = rt.getColumnHeadings(); this code didn't work.
	//String[] ColHead_in_array = ColHead_string.split("\\s*,\\s*"); --- it may have just been the splitting that didn't work.
	int ColN = Array.getLength(ColHead_in_array);
	//int ColN = 9;
	//incase you cannot retrieve the number of columns, assign a variable as above.
	int RowN = ArrayL / ColN;
	//something tells me its better to change int to double for calculations generally.
	//above it shouldn't matter.
	
	///String data_in_matrix[][] = new String[RowN][ColN];				//removed to allow '_' substitution in v1.0.3
	///for(int i=0;i<RowN;i++)											//removed to allow '_' substitution in v1.0.3
	///	   for(int j=0;j<ColN;j++)										//removed to allow '_' substitution in v1.0.3
	///	       data_in_matrix[i][j] = Results_in_array[j%ColN+i*ColN];	//removed to allow '_' substitution in v1.0.3
	//line below is test print out of the data in the 2D array
	///System.out.println(Arrays.deepToString(data_in_matrix));
	
	*/
	
	String currentUsersHomeDir = System.getProperty("user.home");
	String File_Name = (currentUsersHomeDir + File.separator + "Desktop" + File.separator + "Rename me after writing is done.xlsx");
	///String Temp_Write_File = (currentUsersHomeDir + File.separator + "Desktop" + File.separator + "Temporary File to Write data to.xlsx");
	File ExcelFile = new File(File_Name);
	///File ReadFile = new File(Temp_Write_File);		//If creating workbook from file input
	if (!ExcelFile.exists()) {
	OutputStream fileOut = null;
	try {
		fileOut = new FileOutputStream(ExcelFile);
		@SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook();			//non-streamable xlsx workbook
		//SXSSFWorkbook wb = new SXSSFWorkbook();
		wb.createSheet("A");
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fileOut != null)
					fileOut.close();
			}catch (IOException e) {//closing quietly}
				}
			}
	}
	
	/*
	//Try duplicating the original file (or newly created file). Give this duplicate file a name indicating it is temporary. Create the wb object using this file.
	if (ExcelFile.exists()) {
		try {
			copyFile(ExcelFile, ReadFile);
		} catch (IOException e2) {
			e2.printStackTrace();
		}
	}
	*/
	
	InputStream inp = null;
	OutputStream fileOut2 = null;
	try {
	inp = new FileInputStream(ExcelFile);		//Reading file as an inputstream is more memory intensive. Could have tried a bufferedinputstream also.
    ///Workbook wb = WorkbookFactory.create(inp);			//As above. Workbook created from inputstream.
	XSSFWorkbook wb = new XSSFWorkbook(inp);			//previously WorkbookFactory wb = WorkbookFactory.create(ReadFile);
	int Total_sheets = wb.getNumberOfSheets();
    int Last_sheet_ref = Total_sheets - 1;
    Sheet sheet = wb.getSheetAt(Last_sheet_ref);		//better than getSheet("A","B","whatever") as this will always pick the last sheet.
    if (sheet == null)
    	sheet = wb.createSheet("A");
    	//if no sheets exist, then the first will be created as "A".
    Row row = sheet.getRow(2);
    if (row == null)
    	row = sheet.createRow(2);
    	//as for sheet above, but now for the desired row.
    int cell_index = row.getLastCellNum() + 2;
    Cell cell = row.getCell(cell_index);
    if (cell == null)
        cell = row.createCell(cell_index);
    // cell.setCellType(CellType.NUMERIC);		//may not be necessary to set this CellType
    //cell.setCellValue(0);
    
    ///int RowZ = RowN + 2;			//previous reference from v1.0.2
    int RowZ = RTcount + 2;
    ///int ColZ = cell_index + ColN;	//previous reference from v1.0.2
    int ColZ = cell_index + ColHedLen;
    //ColZ is the dynamic reference for the last cell(column) to be written to.
    int RowZ10 = (int) Math.rint(RowZ/10);
    counter = 0;							//for 1D array
    for(int i=2;i<RowZ;i++) {
    	Row rowX = sheet.getRow(i);
    	//rowX is the row being modified in each loop.
    	if (rowX == null) { rowX = sheet.createRow(i);}
    		for(int j=cell_index;j<ColZ;j++) {
    			Cell cellX = rowX.getCell(j);
    			if (cellX == null) { cellX = rowX.createCell(j);}
    			///int I = i - 2;
    			///int J = j - cell_index;
    			///String cell_valueX = data_in_matrix[J][I];			//swapped around I + J from v1.0.2 to v1.0.3
    			///String cell_valueX = data_in_list.get(counter);	//if using ArrayList... possible performance gain by storing and accessing data in 1D rather than 2D array
    			String cell_valueX = data_in_1Darray[counter];	//if using 1D array
    				counter++;									//if using 1D array
    			cellX.setCellValue(cell_valueX);
    			if (cell_valueX.matches(".*[A-Za-z].*") == true){
    				//cellX.setCellType(CellType.STRING);		//unnecessary to define here
    				cellX.setCellValue(cell_valueX);
    			}
    			else {
    				cellX.setCellType(CellType.NUMERIC);
    				Double cell_valueX_num = Double.parseDouble(cell_valueX);
    				cellX.setCellValue(cell_valueX_num);
    			}
    			IJ.showProgress(i, RowZ + (RowZ10));
    		}
    }
    
    /*
    //create a bold font 'style'
    CellStyle cs = wb.createCellStyle();
    Font f = wb.createFont();
    f.setColor((short)Font.COLOR_NORMAL);
    f.setBold(true);
    cs.setFont(f);
    */
    
    //below is code to write headings to the appropriate columns
    Row Head_row = sheet.getRow(1);
    if (Head_row == null)
    	Head_row = sheet.createRow(1);
    	//as for sheet above, but now for the desired row.
    int Head_index = Head_row.getLastCellNum() + 2;
    ///int HeadZ = Head_index + ColN;					//v1.0.2
    int HeadZ = Head_index + ColHeadings.length;
    for (int j=Head_index;j<HeadZ;j++) {
    	Cell cellX = Head_row.getCell(j);
    	if (cellX == null) { cellX = Head_row.createCell(j);
    	int J = j - Head_index;
    	///String cell_headingX = ColHead_in_array[J];		//v1.0.2
    	String cell_headingX = ColHeadings[J];
    	cellX.setCellValue(cell_headingX);
    	//below is applying the bold 'style' to the new cells
    	//cellX.setCellStyle(cs);										//Try without setting style
    	}
    }
    //below... make a new cell prior to headings named 'Count'
    int Count_index = Head_index - 1;
    Cell Count_cell = Head_row.getCell(Count_index);
	if (Count_cell == null) { Count_cell = Head_row.createCell(Count_index);
	Count_cell.setCellValue("Count");
	//Count_cell.setCellStyle(cs);										//Try without setting style
	}
	
	//below... add image title to a cell above the data and column headings
	Row Title_row = sheet.getRow(0);
	if (Title_row == null)
    	Title_row = sheet.createRow(0);
    int Title_index = Head_index - 1;
    Cell Title_cell = Title_row.getCell(Title_index);
	if (Title_cell == null) { Title_cell = Title_row.createCell(Title_index);
		if  (image_title == null) {
	String[] image_name = ij.WindowManager.getImageTitles();
			if (image_name.length == 0) {
				image_title = "";
			}
			else {
				image_title = image_name[0];
			}
		}
	Title_cell.setCellValue("" + image_title);
	//Title_cell.setCellStyle(cs);										//Try without setting style
	}
    
	//below... create column of counts underneath Count_cell up until the last row of data(RowZ).
	for(int i=2;i<RowZ;i++) {
    	Row rowX = sheet.getRow(i);
    	//rowX is the row being modified in each loop.
    	if (rowX == null) { rowX = sheet.createRow(i);}
    	//int CellZ = Total_index; //again for when counts need to start from below Total.
    	int CellZ = Count_index;
    	Cell CellX = rowX.getCell(CellZ);
    	if (CellX == null) { CellX = rowX.createCell(CellZ);}
    	int I = i - 1;
    	CellX.setCellValue(I);
    }
	
	//below... set an underneath border style for the first row of data, separating the Total row from the counts.
	///CellStyle cs2 = wb.createCellStyle();
	///cs2.setBorderBottom(CellStyle.BORDER_THIN);
	///for(int j=cell_index;j<ColZ;j++) {
	///	Row rowX = sheet.getRow(2);
	///	if (rowX == null) { rowX = sheet.createRow(2);}
	///	Cell CellX = rowX.getCell(j);
	///	if (CellX == null) { CellX = rowX.createCell(j);}
	///	CellX.setCellStyle(cs2);
	///}

	// Write the output to a file
	fileOut2 = new FileOutputStream(ExcelFile);
	///int bufferSize = 2 * 1024;
	///BufferedOutputStream bufFileOut = new BufferedOutputStream(fileOut, bufferSize);
	///SXSSFWorkbook wbss = new SXSSFWorkbook(wb, 100);
	wb.write(fileOut2);							///bottleneck might be here
	//bufFileOut.flush();
	//bufFileOut.close();
	fileOut2.flush();
	fileOut2.close();
	wb.close();
	///wbss.dispose();						// dispose of temporary files backing this workbook on disk
	
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}  finally {
		try {
			if (inp != null)
				inp.close();
		}catch (IOException e3){//closing quietly}				
		try {
			if (fileOut2 != null)
				fileOut2.close();
		}catch (IOException e4){//closing quietly}
			}
		}
	}
	
	/*
	if(ReadFile.delete())
    {
        //IJ.log("ReadFile deleted successfully");
    }
    else
    {
    	IJ.log("Failed to delete the temporary file");
    }
	*/
	
	IJ.showProgress(RTcount + 2, RTcount + 2);
	System.gc();
}


/*
@SuppressWarnings("resource")
private static void copyFile(File sourceFile, File destFile) throws IOException {
    if(!destFile.exists()) {
        destFile.createNewFile();
    }

    FileChannel source = null;
    FileChannel destination = null;

    try {
        source = new FileInputStream(sourceFile).getChannel();
        destination = new FileOutputStream(destFile).getChannel();
        destination.transferFrom(source, 0, source.size());
    }
    finally {
        if(source != null) {
            source.close();
        }
        if(destination != null) {
            destination.close();
        }
    }
}
*/

}