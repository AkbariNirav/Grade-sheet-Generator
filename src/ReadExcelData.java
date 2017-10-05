 /*
 * Imported external jxl library for fetching data from excel
 */
import jxl.*;    // Imported complete library
import jxl.read.biff.BiffException;     // Imported inbuit exception for checking the excel format

/*
* Imported java inbuilt library for input/output
*/
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

/*
* Imported java inbuilt library for ArrayList
*/
import java.util.ArrayList;
 
/*
 * The class ReadExcelData firstly fetches the data from the excel sheet which user sets the
 * path and after that it has an instance of class PDFGenerator hence it converts the fetched 
 * data to individual PDF. 
 */
public class ReadExcelData 
{
	
	String filepath = null; // For inputing the file path of excel file
    int numSubjects;        // For inputing the number of subjects
    boolean checker = true;  // For ending the do while loop
     
    public void excel_pdf()   // Method for converting fetched excel data to PDF
	{   
        do   // Start of do-while loop to continue 
        {
	    	try 
	    	{
	    		
	    		// BufferedReader class to store input temporarily
	        	BufferedReader br = new BufferedReader(new InputStreamReader(System.in)); 
		        	
		        System.out.print("Input File Path For Excel :"); // For input of excel file path
		        filepath = br.readLine();
			    
				System.out.println("Enter number of subjects: ");  // For input of subjects
			    numSubjects = Integer.parseInt(br.readLine());
			       	
			   	/*
			    * EXCEL PROCESSING STARTS FROM HERE
			    * 
			    * its (column , row) for indexing
			    */
			        	
			    //Create a workbook object from the file at specified location.
			    //Change the path of the file as per the location on your computer.
			    Workbook wrk1 =  Workbook.getWorkbook(new File(filepath));
			
			    //Obtain the reference to the first sheet in the workbook
			    Sheet sheet1 = wrk1.getSheet(0);
			    			              
			    //getting the name of subjects from excel and saving it to arraylist
			    ArrayList<String> subjNames = new ArrayList<>();
				
			    for(int col = 2; col < numSubjects + 2 ; col++) // For loop to fetch subject input from excel
			    {
			    	Cell subName;    // Refrence of cell inbuilt in jxl
			    	subName = sheet1.getCell(col,0);   // Passing refrence of sheet
			    	subjNames.add(subName.getContents());  // Adding data to arraylist 
			    }
			
			    for(String s : subjNames)  // For loop for printing all subjects
			    {
			    	System.out.println(s);
			    }
			            
			    ArrayList<Student> studentList = new ArrayList<>();  // Arraylist of type Student class
			            
			    for(int row = 1;;row++)   // For loop for fetching student data
			    {
			       	Student student  = new Student();  // Instance of student class calls constructor
			       	
			      	Cell stName;  // Refrence of Cell inbuilt in jxl
			       	try
			       	{
			       		stName = sheet1.getCell(0, row);     //Passing instance of sheet and getting the names
			       	}
			       	catch(ArrayIndexOutOfBoundsException ae)
			       	{
			       		//break the loop when exception occurs, as this is null condition or blank cell
			       		break;
			       	}
			            	
			       	student.name = stName.getContents(); // Passing content to name in Student class
			
			       	Cell stRollNum = sheet1.getCell(1, row);  // Fetching Roll No data from sheet
			      	student.rollNo = stRollNum.getContents();  // Passing content of roll no in Student class
			            	
			       	for(int j = 2;j < numSubjects + 2 ; j++)   // Fetches the subject grade pointers
			       	{
			       		Cell subj = sheet1.getCell(j, row); // Fetches each student data
			       		
			       		// Saves the fetched data into the HashMap of Student class subMarksMap
			      		student.subMarksMap.put(subjNames.get(j-2), Float.parseFloat(subj.getContents()));
			       	}
			            	
			       	studentList.add(student); // Adding to list
			    }
			    
			     //Generate PDF using the list
			     PDFGenerator pdfGen = new PDFGenerator(); // Instance of PDFGenerator class
		         pdfGen.createAllPDFs(studentList); // Passing ArrayList to method of PDFGenerator class
	             
				 checker = false;  // For ending the do-while loop
	        } 
			catch (BiffException e)  // Exception for checking the excel file format
	    	{
				checker = true;  // For continuing the do-while loop
			    System.out.print("Please Check your excel file's format\n\n");  
	    	}
	    	catch (IOException e) // Exception for checking the path of excel file
	    	{
	            checker = true;  // For continuing the do-while loop
	            System.out.print("Please Enter Correct Path\n\n");
	        }
        }
        while(checker);   // End of do-while loop
    } 
}
