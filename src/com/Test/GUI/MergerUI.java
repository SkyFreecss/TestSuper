package com.Test.GUI;

import java.awt.Container;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;

import com.Test.MergerReadExcel;

/**
 * ������
 * @author SkyFreecss
 *
 */
@SuppressWarnings("unused")
public class MergerUI implements ActionListener{
	   
	   List<String> list;
	   File filename;
       JFrame frame = new JFrame("Merger AlphaTest");//��ܲ���
       JTabbedPane tabPane = new JTabbedPane();//ѡ�����
       Container con = new Container();
       JLabel label1 = new JLabel("ѡ��Ŀ¼");
       JLabel label2 = new JLabel("�ļ����Ŀ¼");
       JTextField text1 = new JTextField();//Ŀ¼��·��
       JTextField text2 = new JTextField();//�洢��·��
       JButton button1 = new JButton("...");//ѡ��
       JFileChooser jfc = new JFileChooser();//�ļ�ѡ����
       JButton button2 = new JButton("ȷ��");
       JButton button3 = new JButton("...");//ѡ��
       
       MergerUI()
       {
    	   jfc.setCurrentDirectory(new File("F://"));//�ļ�ѡ�����ĳ�ʼ·��
    	   
    	   double lx = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
    	   double ly = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
    	   
    	   frame.setLocation(new Point((int)(lx/2)-150,(int)(ly/2)-150));//�趨���ڳ���λ��
    	   frame.setSize(280,200);//�趨���ڴ�С
    	   frame.setContentPane(tabPane);//���ò���
    	   label1.setBounds(14,10,70,20);
    	   label2.setBounds(5,35,105,20);
    	   text1.setBounds(100,10,100,20);
    	   text2.setBounds(100,35,100,20);
    	   button1.setBounds(210,10,50,20);
    	   button2.setBounds(30,60,60,20);
    	   button3.setBounds(210,35,50,20);
    	   button1.addActionListener(this);//����¼�����
    	   button2.addActionListener(this);//����¼�����
    	   button3.addActionListener(this);//����¼�����
    	   con.add(label1);
    	   con.add(label2);
    	   con.add(text1);
    	   con.add(text2);
    	   con.add(button1);
    	   con.add(button2);
    	   con.add(button3);
    	   frame.setVisible(true);//���ڿɼ���
    	   frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);//��ʹ���ڹرգ���������
    	   tabPane.add("���1",con);//����1
       }

	public void actionPerformed(ActionEvent e) {
		MergerReadExcel mre = new MergerReadExcel();
           if(e.getSource().equals(button1))//�жϴ��������İ�ť���ĸ�
           {
        	   jfc.setFileSelectionMode(1);//�趨ֻ��ѡ���ļ���
        	   int state = jfc.showOpenDialog(null);//�˾��Ǵ��ļ�ѡ�����Ĵ������
        	   if(state==1)
        	   {
        		   return;
        	   }
        	   else
        	   {
        		   File file = jfc.getSelectedFile();//fileΪѡ���Ŀ¼
        		   text1.setText(file.getAbsolutePath());
        		   list = mre.printFilePath(".xls",file);
        	   }
           }
           
           if(e.getSource().equals(button3))
           {
        	   jfc.setFileSelectionMode(1);//�趨ֻ��ѡ���ļ���
        	   int state = jfc.showOpenDialog(null);
        	   if(state==1)
        	   {
        		   return;
        	   }
        	   else
        	   {
        		   File file = jfc.getSelectedFile();
        		   text2.setText(file.getAbsolutePath());
        		   filename = file;
        	   }
           }
           
           if(e.getSource().equals(button2))
           {
        	   mre.readandwriteExcel(list,filename);
        	   
           }
		
	}
	
	public static void main(String args[])
	{
		new MergerUI();
	}
}
