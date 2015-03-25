/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.yu;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * A simple WOrdprocessingML document created by POI XWPF API
 *
 * @author Yegor Kozlov
 */
public class MakeW {
	private static  List<String> fls=new ArrayList<String>();

	static JFrame jf=new JFrame("cc");
	static Readdoc readdoc;
	static JButton jb=new JButton("click me!");
	static String path;
    public static void main(String[] args) throws Exception {
        /*XWPFDocument doc = new XWPFDocument();

        XWPFParagraph p1 = doc.createParagraph();
        p1.setAlignment(ParagraphAlignment.CENTER);
        p1.setBorderBottom(Borders.DOUBLE);
        p1.setBorderTop(Borders.DOUBLE);

        p1.setBorderRight(Borders.DOUBLE);
        p1.setBorderLeft(Borders.DOUBLE);
        p1.setBorderBetween(Borders.SINGLE);

        p1.setVerticalAlignment(TextAlignment.TOP);

        XWPFRun r1 = p1.createRun();
        r1.setBold(true);
        r1.setText("The quick brown fox");
        r1.setBold(true);
        r1.setFontFamily("Courier");
        r1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        r1.setTextPosition(100);

        XWPFParagraph p2 = doc.createParagraph();
        p2.setAlignment(ParagraphAlignment.RIGHT);

        p2.setBorderBottom(Borders.DOUBLE);
        p2.setBorderTop(Borders.DOUBLE);
        p2.setBorderRight(Borders.DOUBLE);
        p2.setBorderLeft(Borders.DOUBLE);
        p2.setBorderBetween(Borders.SINGLE);

        XWPFRun r2 = p2.createRun();
        r2.setText("jumped over the lazy dog");
        r2.setStrike(true);
        r2.setFontSize(20);

        XWPFRun r3 = p2.createRun();
        r3.setText("and went away");
        r3.setStrike(true);
        r3.setFontSize(20);
        r3.setSubscript(VerticalAlign.SUPERSCRIPT);


        XWPFParagraph p3 = doc.createParagraph();
        p3.setWordWrap(true);
        p3.setPageBreak(true);
                
        p3.setAlignment(ParagraphAlignment.BOTH);
        p3.setSpacingLineRule(LineSpacingRule.EXACT);

        p3.setIndentationFirstLine(600);
        

        XWPFRun r4 = p3.createRun();
        r4.setTextPosition(20);
        r4.setText("To be, or not to be: that is the question: "
                + "Whether 'tis nobler in the mind to suffer "
                + "The slings and arrows of outrageous fortune, "
                + "Or to take arms against a sea of troubles, "
                + "And by opposing end them? To die: to sleep; ");
        r4.addBreak(BreakType.PAGE);
        r4.setText("No more; and by a sleep to say we end "
                + "The heart-ache and the thousand natural shocks "
                + "That flesh is heir to, 'tis a consummation "
                + "Devoutly to be wish'd. To die, to sleep; "
                + "To sleep: perchance to dream: ay, there's the rub; "
                + ".......");
        r4.setItalic(true);

        XWPFRun r5 = p3.createRun();
        r5.setTextPosition(-10);
        r5.setText("For in that sleep of death what dreams may come");
        r5.addCarriageReturn();
        r5.setText("When we have shuffled off this mortal coil,"
                + "Must give us pause: there's the respect"
                + "That makes calamity of so long life;");
        r5.addBreak();
        r5.setText("For who would bear the whips and scorns of time,"
                + "The oppressor's wrong, the proud man's contumely,");
        
        r5.addBreak(BreakClear.ALL);
        r5.setText("The pangs of despised love, the law's delay,"
                + "The insolence of office and the spurns" + ".......");

        FileOutputStream out = new FileOutputStream("simple.docx");
        doc.write(out);
        out.close();*/

        /*Writedoc wdoc=new Writedoc();
        wdoc.testWriteTable();*/
    	//search pwd docs
    	
        /*Readdoc rdoc=new Readdoc();
        rdoc.testReadByDoc();*/
        /*Replacedoc redoc=new Replacedoc();
        redoc.Replacedoc();*/
    	jf.setSize(100, 100);
    	jf.add(jb);
    	
    	//final JFileChooser filechooser = new JFileChooser();//创建文件选择器
		
    	jb.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				 /*FileDialog fd=new FileDialog(new Frame(),"测试",FileDialog.LOAD);
				 FilenameFilter ff=new FilenameFilter(){
				   public boolean accept(File dir, String name) {
				    if (name.endsWith("docx")){
				     return true;
				    }
				    return false;
				   }
				  };
				  fd.setFilenameFilter(ff);
				 fd.setVisible(true);
				 System.out.println(fd.getDirectory()+fd.getFile());*/
				// JFileChooser filechooser = new JFileChooser();//创建文件选择器
				    /*filechooser.setCurrentDirectory(new File("."));//设置当前目录
				    filechooser.setAcceptAllFileFilterUsed(false);
				    //显示所有文件
				    filechooser.addChoosableFileFilter(new javax.swing.filechooser.FileFilter() {
				      public boolean accept(File f) {
				        return true;
				      }
				      public String getDescription() {
				        return "所有文件(*.*)";
				      }
				    });
				    //显示JAVA源文件
				    filechooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
				      public boolean accept(File f) { //设定可用的文件的后缀名
				        if(f.getName().endsWith(".docx")||f.isDirectory()){
				          return true;
				        }
				        return false;
				      }
				      public String getDescription() {
				        return "报告(*.docx)";
				      }
				    });
				    
				    int returnVal = filechooser.showOpenDialog(SimpleDocument.this);
				    if(returnVal == JFileChooser.APPROVE_OPTION) {
				       System.out.println("You chose to open this file: " +
				    		   filechooser.getSelectedFile().getName());
				    }

				 path=filechooser.getSelectedFile().getAbsolutePath();*/
				
			        
				JFileChooser filechooser = new JFileChooser();
		        filechooser.setCurrentDirectory(new File("."));
		        filechooser.setAcceptAllFileFilterUsed(false);//是否显示所有文件
		        FileNameExtensionFilter filter = new FileNameExtensionFilter("WORD file", "doc", "docx");
		        filechooser.addChoosableFileFilter(filter);


		          int  ret = filechooser.showOpenDialog(null);

		        
		        if (ret == JFileChooser.APPROVE_OPTION) {
		             path=filechooser.getSelectedFile().getAbsolutePath();
		        }
				

				//path=fd.getFile();
				System.out.println("path==============="+path);
				 ListFile lf=new ListFile();
			    	fls=lf.listFile(path, "", "doc");
			    	readdoc=new Readdoc();
    	//readdoc.testReadByDoc("/home/yu/workspace/forw/lizi.docx");
    	try {
			readdoc.testReadByDoc(fls);
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
    	try {
			Replacedoc redoc=new Replacedoc(readdoc.params,fls);
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
			    	System.out.println("heare");
				jb.setText("^^");
				JOptionPane.showMessageDialog(null, "Success！", "消息提示", JOptionPane.OK_OPTION);
			} 
    		
    	});
    	/*readdoc=new Readdoc();
    	//readdoc.testReadByDoc("/home/yu/workspace/forw/lizi.docx");
    	readdoc.testReadByDoc(fls);
    	Replacedoc redoc=new Replacedoc(readdoc.params);*/
    	jf.setLocation(300, 300);
    	jf.show();
    	
    	/*ListFile lf=new ListFile();
    	lf.listFile(path,"","doc|docx");
    	fls=lf.listFile("", "", "docx");*/
    	//System.out.println(fls.toString());
    	
    	
    }

}
