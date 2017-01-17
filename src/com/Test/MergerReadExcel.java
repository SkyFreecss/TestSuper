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
 * �����д�����ز���
 * @author SkyFreecss
 *
 */
@SuppressWarnings("deprecation")
public class MergerReadExcel {
       static Log log = LogFactory.getLog("MergerReadExcel.class");
       static int rowsNum = 0;//�ܱ���
       static int rowsNum_2 = 0;//������
       static String reg = "����";
       static int rowsNum_spc;
	   static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");//�������ڸ�ʽ��
       static String date = sdf.format(new Date());
	   
	   /**
	    * ��ȡĿ¼�º�׺Ϊ.xls���ļ�
	    * @param str
	    * @param file
	    * @return
	    */
	   public  List<String> printFilePath(String str,File file)
	   {
		   log.info("�빺�����棡");
		   log.info("������ֻ֧��xls�ļ���лл��");
		   List<String> list = new Vector<String>();
		   if(file!=null)
		   {
			   log.info("���ڴ򿪸�Ŀ¼��");
			   file.isDirectory();
			   File[] fileArray = file.listFiles();
			   if(fileArray!=null)
			   {
				   log.info("���ڻ�ȡ�����ļ���");
				   for(int i = 0;i<fileArray.length;i++)
				   {
					   
					   if(!fileArray[i].isDirectory())
					   {
						   String tempName = fileArray[i].getName();
						   //�ж��Ƿ�Ϊ.xls��β��
						   if(tempName.trim().toLowerCase().endsWith(str))
						   {
							   System.out.println(fileArray[i]);
							   String fileName = fileArray[i].toString();
							   list.add(fileName);
						   }
					   }
				   }
				   log.info("�ļ���ȡ��ɣ�");
			   }
		   }
		   else
		   {
			   log.error("·������");
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
			
			//-------------------------�ܱ�-------------------------------
			WritableSheet ws = wwb.createSheet("�ܱ�",0);
			
			WritableCellFormat format = new WritableCellFormat();
			format.setAlignment(Alignment.CENTRE);//���þ��С�
			format.setVerticalAlignment(VerticalAlignment.CENTRE);//ͬ�ϡ�
			format.setBackground(Colour.SKY_BLUE);//����ɫ
			format.setWrap(true);//�Զ�����
			
		    ws.setColumnView(0,10);//��Ԫ���С
		    ws.setColumnView(1,10);
		    ws.setColumnView(2,30);
		    ws.setColumnView(3,7);
		    ws.setColumnView(4,80);
		    ws.setColumnView(5,20);
		    ws.setColumnView(6,10);
		    ws.setColumnView(7,10);
			ws.setColumnView(8,40);
		    
			//�½���ĵ�һ�У���ͷ��
			Label label = new Label(0,0,"������ҵ��",format);
			Label label_1 = new Label(1,0,"�漰��Ŀ������",format);
			Label label_2 = new Label(2,0,"ģ�����������",format);
			Label label_3 = new Label(3,0,"������Ա",format);
			Label label_4 = new Label(4,0,"���幤������",format);
			Label label_5 = new Label(5,0,"�ƻ���������",format);
			Label label_6 = new Label(6,0,"ʵ���������",format);
			Label label_7 = new Label(7,0,"������",format);
			Label label_8 = new Label(8,0,"��ע",format);
			
			ws.addCell(label);
			ws.addCell(label_1);
			ws.addCell(label_2);
			ws.addCell(label_3);
			ws.addCell(label_4);
			ws.addCell(label_5);
			ws.addCell(label_6);
			ws.addCell(label_7);
			ws.addCell(label_8);
			
			//-------------------------�ܱ�-------------------------------
			WritableSheet ws2 = wwb.createSheet("����",1);
			
			ws2.setColumnView(0,10);//��Ԫ���С
		    ws2.setColumnView(1,10);
		    ws2.setColumnView(2,10);
		    ws2.setColumnView(3,10);
		    ws2.setColumnView(4,80);
		    ws2.setColumnView(5,30);
		    ws2.setColumnView(6,10);
		    ws2.setColumnView(7,10);
		    ws2.setColumnView(8,50);
		    
			Label label2_1 = new Label(0,0,"������ҵ��",format);
			Label label2_2 = new Label(1,0,"�漰��Ŀ������",format);
			Label label2_3 = new Label(2,0,"ģ�����������",format);
			Label label2_4 = new Label(3,0,"������Ա",format);
			Label label2_5 = new Label(4,0,"���幤������",format);
			Label label2_6 = new Label(5,0,"�ƻ���������",format);
			Label label2_7 = new Label(6,0,"ʵ���������",format);
			Label label2_8 = new Label(7,0,"������",format);
			Label label2_9 = new Label(8,0,"��ע",format);
			
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
			log.info("�ѻ�ȡ��·����Ϣ�����ڽ��б�Ļ�ȡ��");
			
			
			//�Ա��ļ����б���
			for(int i=0;i<list.size();i++)
			{
				String filename = list.get(i);
				log.info("�ѻ�ȡ��"+filename);
				System.out.println(filename);
				//��������������
				InputStream is = new FileInputStream(filename);
				
				//��������������
				Workbook wb = Workbook.getWorkbook(is);
				
				//��ù������ĸ���
				wb.getNumberOfSheets();
				
				Sheet oFirstSheet = wb.getSheet(0);//ʹ����������ʽ��ȡ��һ��������
				
				
				if(filename.indexOf(reg)!=-1)
				{
					System.out.println(filename);
					int rows = oFirstSheet.getRows();//��ȡ���������
					int columns = oFirstSheet.getColumns();//��ȡ���������
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
					int rows = oFirstSheet.getRows();//��ȡ���������
					int columns = oFirstSheet.getColumns();//��ȡ���������
				for(int x = rowsNum;x<rows+rowsNum;x++)
				{
					for(int y = 0;y<columns;y++)
					{
						rowsNum_spc = x-rowsNum;
						if(rowsNum_spc!=0)
						{
						Cell ocell2 = oFirstSheet.getCell(y,rowsNum_spc);//��ȡ��Ԫ������
						Label label2 = new Label(y,x,ocell2.getContents());//��һ��������Ҫд����������ڶ�����д���������������Ǵ����������
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
		   log.info("�ѻ�ȡ��·����Ϣ�����ڽ��б�Ļ�ȡ��");
		   for(int i=0;i<list.size();i++)
		   {
			   String filename = list.get(i);
			   log.info("�ѻ�ȡ��"+filename);
			   System.out.println(filename);
				//��������������
				InputStream is = new FileInputStream(filename);
				
				//��������������
				Workbook wb = Workbook.getWorkbook(is);
				
				//��ù������ĸ���
				wb.getNumberOfSheets();
				
				Sheet oFirstSheet = wb.getSheet(0);//ʹ����������ʽ��ȡ��һ��������
				
				int rows = oFirstSheet.getRows();//��ȡ����������������
				int columns = oFirstSheet.getColumns();//��ȡ����������������
				
				log.info("���������"+filename+" ������");
				for(int x = 0;x<rows;x++)
				{
					for(int y = 0;y<columns;y++)
					{
						Cell ocell = oFirstSheet.getCell(y, x);//������������һ�����������ڶ�����������
						System.out.println(ocell.getContents());
					}
				}
				log.info(filename+"���������ɣ�");
			} catch (BiffException | IOException e) {
				log.error("�������⣡");
				e.printStackTrace();
			}
		   }
		   log.info("���б��ѻ�ȡ��ɣ�");
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
