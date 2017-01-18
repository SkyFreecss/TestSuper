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
 * 界面类
 * @author SkyFreecss
 *
 */
@SuppressWarnings("unused")
public class MergerUI implements ActionListener{
	   
	   List<String> list;
	   File filename;
       JFrame frame = new JFrame("Merger AlphaTest");//框架布局
       JTabbedPane tabPane = new JTabbedPane();//选项卡布局
       Container con = new Container();
       JLabel label1 = new JLabel("选择目录");
       JLabel label2 = new JLabel("文件存放目录");
       JTextField text1 = new JTextField();//目录的路径
       JTextField text2 = new JTextField();//存储的路径
       JButton button1 = new JButton("...");//选择
       JFileChooser jfc = new JFileChooser();//文件选择器
       JButton button2 = new JButton("确定");
       JButton button3 = new JButton("...");//选择
       
       MergerUI()
       {
    	   jfc.setCurrentDirectory(new File("F://"));//文件选择器的初始路径
    	   
    	   double lx = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
    	   double ly = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
    	   
    	   frame.setLocation(new Point((int)(lx/2)-150,(int)(ly/2)-150));//设定窗口出现位置
    	   frame.setSize(280,200);//设定窗口大小
    	   frame.setContentPane(tabPane);//设置布局
    	   label1.setBounds(14,10,70,20);
    	   label2.setBounds(5,35,105,20);
    	   text1.setBounds(100,10,100,20);
    	   text2.setBounds(100,35,100,20);
    	   button1.setBounds(210,10,50,20);
    	   button2.setBounds(30,60,60,20);
    	   button3.setBounds(210,35,50,20);
    	   button1.addActionListener(this);//添加事件处理
    	   button2.addActionListener(this);//添加事件处理
    	   button3.addActionListener(this);//添加事件处理
    	   con.add(label1);
    	   con.add(label2);
    	   con.add(text1);
    	   con.add(text2);
    	   con.add(button1);
    	   con.add(button2);
    	   con.add(button3);
    	   frame.setVisible(true);//窗口可见度
    	   frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);//能使窗口关闭，结束程序
    	   tabPane.add("面板1",con);//布局1
       }

	public void actionPerformed(ActionEvent e) {
		MergerReadExcel mre = new MergerReadExcel();
           if(e.getSource().equals(button1))//判断触发方法的按钮是哪个
           {
        	   jfc.setFileSelectionMode(1);//设定只能选择文件夹
        	   int state = jfc.showOpenDialog(null);//此句是打开文件选择器的触发语句
        	   if(state==1)
        	   {
        		   return;
        	   }
        	   else
        	   {
        		   File file = jfc.getSelectedFile();//file为选择的目录
        		   text1.setText(file.getAbsolutePath());
        		   list = mre.printFilePath(".xls",file);
        	   }
           }
           
           if(e.getSource().equals(button3))
           {
        	   jfc.setFileSelectionMode(1);//设定只能选择文件夹
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
