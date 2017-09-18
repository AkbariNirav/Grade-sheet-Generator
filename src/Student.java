import java.util.HashMap;             // Imported java inbuilt package for using HashMap 

/*
 * Student class to store all the name and roll number of students 
 * which are fetched from the excel file and are store in an HashMap.
 */

public class Student				// Public class Student Declaration 
{
	
	public String name;             // String to store the name of student 
	public String rollNo;           // String to store the roll no of student
	public HashMap<String, Float> subMarksMap;           // Declaration of HashMap

	public Student()          // Default constructor of the class
	{
		subMarksMap = new HashMap<>();    // Assigning memory to the HashMap
	}
}
