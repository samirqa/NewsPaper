import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Script 
{
	    static File file;
	    static FileInputStream inputStream;
	    static XSSFWorkbook WB = null;
	    static XSSFSheet Sheet;
	    static XSSFRow row;
 	    static int numberOfDays = 0;
 	    static int numberWeekEndDays = 0;
 	    static int  numberofSund = 0;
 	    static int  numberofSat = 0;
		public static void main(String args[])throws Exception
		{
			String FilePath = "D:\\eclipse\\workspace\\demoproj\\NewPaper\\SM\\TD.xlsx";
			CallExcel(FilePath,"TD");
			System.out.println("TOTAL NO OF ROWS : "+lastrowb());
			for(int p=1;p<=lastrowb();p++)
			{
				String NewsPaper = getdata(p,0);
				String SDate = getdata(p,1);
				Date date1=new SimpleDateFormat("dd/MM/yyyy").parse(SDate); 
				String EDate = getdata(p,2);
				Date date2=new SimpleDateFormat("dd/MM/yyyy").parse(EDate); 
				String SubType = getdata(p,3);
				String DValid = getdata(p,4);
				System.out.println("News Paper : "+NewsPaper);
				System.out.println("Start Date : "+SDate);
				System.out.println("End Date : "+EDate);
				System.out.println("Subscription Type : "+SubType);
				System.out.println("Date Validation : "+DValid);
				//Date Validation condition
				if(DValid.equals("true"))
				{
					List<String> plist = new ArrayList<String>();
					
					CallExcel(FilePath,"DataSet");			
					//Fetch Prices from Newspaper Data
					for(int i = 1;i<=lastrow();i++)
					{
						if(getdata(i, 0).equals(NewsPaper))
						{
							int y=1;
							for(;y<=7;y++)
							{
								plist.add(getdata(i, y));
							}
						}
					}
					System.out.println("SIZE : "+plist.size());
					for(int y=0;y<plist.size();y++)
					{
						System.out.println(plist.get(y));
					}	
					getWorkingDaysName(date1, date2);
					
					//Subscription Type Logic
					if(SubType.equals("BiWeekly"))
					{
						double totalSatPrice = numberofSat* Double.parseDouble(plist.get(5));
						double totalSunPrice = numberofSund*Double.parseDouble(plist.get(6));
						System.out.println("Total Subscribtion Price  : "+(totalSatPrice+totalSunPrice));
						CallExcel(FilePath,"TD");
						setdata(p, 5, ""+(totalSatPrice+totalSunPrice));
					}
					else if(SubType.equals("Weekly"))
					{
						double totalSatPrice = numberofSat* Double.parseDouble(plist.get(5));
						double totalSunPrice = numberofSund*Double.parseDouble(plist.get(6));
						int totaldays = numberOfDays-numberWeekEndDays;
						double totalWeeklySubscribtionPrice = totaldays*Double.parseDouble(plist.get(1));
						
						System.out.println("Total Weekly Subscribtion Price  : "+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
						CallExcel(FilePath,"TD");
						setdata(p, 5, ""+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
					}
					else if(SubType.equals("Monthly"))
					{
						if(numberOfDays>=7 && numberOfDays<=31)
						{
							double totalSatPrice = numberofSat* Double.parseDouble(plist.get(5));
							double totalSunPrice = numberofSund*Double.parseDouble(plist.get(6));
							int totaldays = numberOfDays-numberWeekEndDays;
							double totalWeeklySubscribtionPrice = totaldays*Double.parseDouble(plist.get(1));
							
							
							System.out.println("Total Weekly Subscribtion Price  : "+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
							CallExcel(FilePath,"TD");
							setdata(p, 5, ""+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
						}
						else
						{
							System.out.println("Price Enter Valid Date Range for Monthly Subscription");
						}
					}
					else if(SubType.equals("Yearly"))
					{
						if(numberOfDays>=350 && numberOfDays<=366)
						{

							double totalSatPrice = numberofSat* Double.parseDouble(plist.get(5));
							double totalSunPrice = numberofSund*Double.parseDouble(plist.get(6));
							int totaldays = numberOfDays-numberWeekEndDays;
							double totalWeeklySubscribtionPrice = totaldays*Double.parseDouble(plist.get(1));
							
							System.out.println("Total Weekly Subscribtion Price  : "+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
							CallExcel(FilePath,"TD");
							setdata(p, 5, ""+(totalWeeklySubscribtionPrice+totalSatPrice+totalSunPrice));
						}
						else
						{
							System.out.println("Price Enter Valid Date Range for Yearly Subscription");
						}
					}
					SaveExcel();
				}
				else
				{
					System.out.println("Invalid Date, please enter valid date");
				}
			}
			
		}
		public static void CallExcel(String FullPath,String SheetName) throws Exception
		{
			file = new File(FullPath);
		    inputStream = new FileInputStream(file);
		    WB = new XSSFWorkbook(inputStream);
		    Sheet = WB.getSheet(SheetName);
		}
		public static void SaveExcel() throws Exception
		{   
			inputStream.close();
		    FileOutputStream outputStream = new FileOutputStream(file);
		    WB.write(outputStream);
		    outputStream.close();
		}
		public static String getdata(int Row,int Column) throws Exception
		{
			String data=null;
			row = Sheet.getRow(Row);
			String datatype = ""+row.getCell(Column).getCellTypeEnum();
			
			if(datatype.equals("STRING"))
			{
				data =  row.getCell(Column).getStringCellValue();
			}
			else if(datatype.equals("NUMERIC"))
			{
//				data = ""+row.getCell(Column).getNumericCellValue();
				DataFormatter dataFormatter = new DataFormatter();
				String cellStringValue = dataFormatter.formatCellValue(row.getCell(Column));
				data = cellStringValue;
			}
			else if(datatype.equals("FORMULA"))
			{
				data = ""+row.getCell(Column).getBooleanCellValue();
			}
			return data;
		}
		public static void setdata(int Row,int Column,String data) throws Exception
		{
			row = Sheet.getRow(Row);
			//String datatype = ""+row.getCell(Column).getCellTypeEnum();
			 Cell cellTitle = row.createCell(Column);
			 cellTitle.setCellValue(data);
		}
		public static void CloseExcel()throws Exception
		{
			WB.close();
		}
		public static int lastrow()
		{
			 return Sheet.getLastRowNum();
		}
		public static int lastrowb()
		{
			String data=null;
			int countf =0;
			try{
				for(int y=0;y<=250;y++)
				{
					row = Sheet.getRow(y);
					String datatype = ""+row.getCell(0).getStringCellValue();
					if(!datatype.equals(""))
					{
						countf=countf+1;
					}
					else
					{
						break;
					}
				}
			}catch(NullPointerException e)
			{
				countf=countf-1;
			}
			return countf;
		}
	 	public static void getWorkingDaysName(Date d1,Date d2)throws Exception
	 	{
	 	    Date date11 = d1;
	 	    Date date22 = d2;
	 	    Calendar cal1 = Calendar.getInstance();
	 	    Calendar cal2 = Calendar.getInstance();
	 	    cal1.setTime(date11);
	 	    cal2.setTime(date22);

	 	    numberOfDays = 0;
	 	    numberWeekEndDays = 0;
	 	    numberofSund = 0;
	 	    numberofSat = 0;
	 	    while (cal1.before(cal2)) 
	 	    {
	 	    	if(cal1.get(Calendar.DATE)==1)
	 	    	{
	 	    		numberOfDays=numberOfDays+1;
//	 	    		System.out.println("Day of Week : "+cal1.get(Calendar.DAY_OF_WEEK));
//	 	    		System.out.println("ITERATE DATE : "+cal1.get(Calendar.DATE));
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==7 || cal1.get(Calendar.DAY_OF_WEEK)==1)
	 	    		{
	 	    			numberWeekEndDays=numberWeekEndDays+1;
	 	    		}
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==7)
	 	    		{
	 	    			numberofSat=numberofSat+1;
	 	    		}
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==1)
	 	    		{
	 	    			numberofSund=numberofSund+1;
	 	    		}
	 	    	}
	 	        cal1.add(Calendar.DATE,1);
	 	        if(cal1.get(Calendar.DATE)>=2)
	 	        {
	 	        	numberOfDays=numberOfDays+1;
//	 	        	System.out.println("Day of Week : "+cal1.get(Calendar.DAY_OF_WEEK));
//	 	        	System.out.println("ITERATE DATE : "+cal1.get(Calendar.DATE));
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==7 || cal1.get(Calendar.DAY_OF_WEEK)==1)
	 	    		{
	 	    			numberWeekEndDays=numberWeekEndDays+1;
	 	    		}
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==7)
	 	    		{
	 	    			numberofSat=numberofSat+1;
	 	    		}
	 	    		if(cal1.get(Calendar.DAY_OF_WEEK)==1)
	 	    		{
	 	    			numberofSund=numberofSund+1;
	 	    		}
	 	        }
	 	    }
	 	   System.out.println("TOTAL No OF WeekEnd Days : "+numberWeekEndDays);
	 	   System.out.println("Total Month Days :"+numberOfDays);
	 	   System.out.println("TOTAL No OF Saturdays : "+numberofSat);
	 	   System.out.println("Total Month Sundays :"+numberofSund);
	 	}
}
