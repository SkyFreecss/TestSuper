package com.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Vector;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Alignment;
import jxl.write.Colour;
import jxl.write.Label;
import jxl.write.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * 读表和写表的相关操作
 * @author SkyFreecss
 *
 */
@SuppressWarnings("deprecation")
public class MergerReadExcel {
       static Log log = LogFactory.getLog("MergerReadExcel.class");
       static int rowsNum = 0;//周报用
       static int rowsNum_2 = 0;//需求用
       static String reg = "需求";
       static int rowsNum_spc;
	   static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");//设置日期格式。
       static String date = sdf.format(new Date());
	   
	   /**
	    * 获取目录下后缀为.xls的文件
	    * @param str
	    * @param file
	    * @return
	    */
	   public  List<String> printFilePath(String str,File file)
	   {
		   log.info("请购买正版！");
		   log.info("本工具只支持xls文件，谢谢！");
		   List<String> list = new Vector<String>();
		   if(file!=null)
		   {
			   log.info("正在打开该目录！");
			   file.isDirectory();
			   File[] fileArray = file.listFiles();
			   if(fileArray!=null)
			   {
				   log.info("正在获取所需文件！");
				   for(int i = 0;i<fileArray.length;i++)
				   {
					   
					   if(!fileArray[i].isDirectory())
					   {
						   String tempName = fileArray[i].getName();
						   //判断是否为.xls结尾！
						   if(tempName.trim().toLowerCase().endsWith(str))
						   {
							   System.out.println(fileArray[i]);
							   String fileName = fileArray[i].toString();
							   list.add(fileName);
						   }
					   }
				   }
				   log.info("文件获取完成！");
			   }
		   }
		   else
		   {
			   log.error("路径有误！");
		   }
		   return list;
	   }
	   
	   
	   /**
	    * 
	    * @param list
	    */
	   @SuppressWarnings({ "unused" })
	public void readandwriteExcel(List<String> list)
	   {
		   try {
			WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel//New_Test_"+date+".xls"));
			
			//-------------------------周报-------------------------------
			WritableSheet ws = wwb.createSheet("周报",0);
			
			WritableCellFormat format = new WritableCellFormat();
			format.setAlignment(Alignment.CENTRE);//设置居中。
			format.setVerticalAlignment(VerticalAlignment.CENTRE);//同上。
			format.setBackground(Colour.SKY_BLUE);//背景色
			format.setWrap(true);//自动换行
			
		    ws.setColumnView(0,10);//单元格大小
		    ws.setColumnView(1,10);
		    ws.setColumnView(2,30);
		    ws.setColumnView(3,7);
		    ws.setColumnView(4,80);
		    ws.setColumnView(5,20);
		    ws.setColumnView(6,10);
		    ws.setColumnView(7,10);
			ws.setColumnView(8,40);
		    
			//新建表的第一行（表头）
			Label label = new Label(0,0,"归属事业部",format);
			Label label_1 = new Label(1,0,"涉及项目或需求",format);
			Label label_2 = new Label(2,0,"模块或需求名称",format);
			Label label_3 = new Label(3,0,"参与人员",format);
			Label label_4 = new Label(4,0,"具体工作内容",format);
			Label label_5 = new Label(5,0,"计划工作周期",format);
			Label label_6 = new Label(6,0,"实际完成天数",format);
			Label label_7 = new Label(7,0,"完成情况",format);
			Label label_8 = new Label(8,0,"备注",format);
			
			ws.addCell(label);
			ws.addCell(label_1);
			ws.addCell(label_2);
			ws.addCell(label_3);
			ws.addCell(label_4);
			ws.addCell(label_5);
			ws.addCell(label_6);
			ws.addCell(label_7);
			ws.addCell(label_8);
			
			//-------------------------周报-------------------------------
			WritableSheet ws2 = wwb.createSheet("需求",1);
			
			ws2.setColumnView(0,10);//单元格大小
		    ws2.setColumnView(1,10);
		    ws2.setColumnView(2,10);
		    ws2.setColumnView(3,10);
		    ws2.setColumnView(4,80);
		    ws2.setColumnView(5,30);
		    ws2.setColumnView(6,10);
		    ws2.setColumnView(7,10);
		    ws2.setColumnView(8,50);
		    
			Label label2_1 = new Label(0,0,"归属事业部",format);
			Label label2_2 = new Label(1,0,"涉及项目或需求",format);
			Label label2_3 = new Label(2,0,"模块或需求名称",format);
			Label label2_4 = new Label(3,0,"参与人员",format);
			Label label2_5 = new Label(4,0,"具体工作内容",format);
			Label label2_6 = new Label(5,0,"计划工作周期",format);
			Label label2_7 = new Label(6,0,"实际完成天数",format);
			Label label2_8 = new Label(7,0,"完成情况",format);
			Label label2_9 = new Label(8,0,"备注",format);
			
			ws2.addCell(label2_1);
			ws2.addCell(label2_2);
			ws2.addCell(label2_3);
			ws2.addCell(label2_4);
			ws2.addCell(label2_5);
			ws2.addCell(label2_6);
			ws2.addCell(label2_7);
			ws2.addCell(label2_8);
			ws2.addCell(label2_9);
			
			/*
			Alignment alignment=null;
			VerticalAlignment verticalAlignment=null;
			Colour color=null;
			boolean flag;
			*/
			log.info("已获取到路径信息，正在进行表的获取！");
			
			
			//对表文件进行遍历
			for(int i=0;i<list.size();i++)
			{
				String filename = list.get(i);
				log.info("已获取表："+filename);
				System.out.println(filename);
				//构造输入流对象
				InputStream is = new FileInputStream(filename);
				
				//声明工作薄对象
				Workbook wb = Workbook.getWorkbook(is);
				
				//获得工作薄的个数
				wb.getNumberOfSheets();
				
				Sheet oFirstSheet = wb.getSheet(0);//使用索引的形式获取第一个工作表。
				
				
				if(filename.indexOf(reg)!=-1)
				{
					System.out.println(filename);
					int rows = oFirstSheet.getRows();//获取表的总行数
					int columns = oFirstSheet.getColumns();//获取表的总列数
				   for(int m = rowsNum_2;m<rows+rowsNum_2;m++)
				   {
					   for(int n=0;n<columns;n++)
					   {
						   rowsNum_spc = m-rowsNum_2;
						   if(rowsNum_spc!=0)
						   {
						   Cell ocell1 = oFirstSheet.getCell(n,rowsNum_spc);
						   Label label1 = new Label(n,m,ocell1.getContents(),ocell1.getCellFormat());
						   ws2.addCell(label1);
						   }
					   }
				   }
				}
				else
				{
					int rows = oFirstSheet.getRows();//获取表的总行数
					int columns = oFirstSheet.getColumns();//获取表的总列数
				for(int x = rowsNum;x<rows+rowsNum;x++)
				{
					for(int y = 0;y<columns;y++)
					{
						rowsNum_spc = x-rowsNum;
						if(rowsNum_spc!=0)
						{
						Cell ocell2 = oFirstSheet.getCell(y,rowsNum_spc);//获取单元格内容
						Label label2 = new Label(y,x,ocell2.getContents());//第一个参数的要写入的列数，第二个是写入行数，第三个是此坐标的内容
						ws.addCell(label2);
						}
					}
				}
				rowsNum = rowsNum+rows;	
				}
				
				}
			
			wwb.write();
			wwb.close();
			
			
		} catch (IOException | BiffException | WriteException e) {
			e.printStackTrace();
		}
		finally
		{
			String cmd[] = {"C:\\Program Files (x86)\\OpenOffice 4\\program\\scalc.exe","F\\TestFile\\Excel\\New_Test_"+date+".xls"};
			try {
				Process p = Runtime.getRuntime().exec(cmd);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		 /*
		   WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel//New_Test.xls"));
		   log.info("已获取到路径信息，正在进行表的获取！");
		   for(int i=0;i<list.size();i++)
		   {
			   String filename = list.get(i);
			   log.info("已获取表："+filename);
			   System.out.println(filename);
				//构造输入流对象
				InputStream is = new FileInputStream(filename);
				
				//声明工作薄对象
				Workbook wb = Workbook.getWorkbook(is);
				
				//获得工作薄的个数
				wb.getNumberOfSheets();
				
				Sheet oFirstSheet = wb.getSheet(0);//使用索引的形式获取第一个工作表
				
				int rows = oFirstSheet.getRows();//获取工作薄的总行数。
				int columns = oFirstSheet.getColumns();//获取工作薄的总列数。
				
				log.info("正在输出表："+filename+" 的内容");
				for(int x = 0;x<rows;x++)
				{
					for(int y = 0;y<columns;y++)
					{
						Cell ocell = oFirstSheet.getCell(y, x);//两个参数，第一个是列数，第二个是行数。
						System.out.println(ocell.getContents());
					}
				}
				log.info(filename+"内容输出完成！");
			} catch (BiffException | IOException e) {
				log.error("严重问题！");
				e.printStackTrace();
			}
		   }
		   log.info("所有表都已获取完成！");
		   */
	   }
	   
	   
	   
	   public static void main(String args[])
	   {
		   MergerReadExcel mre = new MergerReadExcel();
		   File file = new File("F://TestFile//Excel//0103-0106");
		   String str=".xls";
		   List<String> list = mre.printFilePath(str, file);
		   mre.readandwriteExcel(list);
	   }
}
