package com.yu;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.text.Keymap;
import javax.xml.crypto.dsig.keyinfo.KeyValue;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

public class Readdoc {
	 
	Map<String, String> params = new HashMap<String, String>();
	   /**
	    * 通过XWPFDocument对内容进行访问。对于XWPF文档而言，用这种方式进行读操作更佳。
	    * @throws Exception
	    */
	   public void testReadByDoc( List<String> fPath) throws Exception {
	      
		 for(String path:fPath){
			 System.out.println("path==="+path);
			 if(!path.contains("模板")){
				 
			   	
		 
		 System.out.println("table mun=");
		 
		 InputStream is = new FileInputStream(path);
	      HWPFDocument doc = new HWPFDocument(is);
	      

	      Range range = doc.getRange();   ////////////////
	      
	      int secNum = range.numSections();  
	     //// System.out.println("secnum="+secNum);  
	      Section section;  
	      for (int i=0; i<secNum; i++) {  
	         section = range.getSection(i);  
	          
	         //System.out.println(section.text());  
	      } 
	      ////////////////////////////////////////////////////////
	         TableIterator tables = new TableIterator(range);  
//////////////////////////////
/*int num = range.numParagraphs();  
System.out.println("num="+num);
Paragraph para;  
for (int i=0; i<num; i++) {  
para = range.getParagraph(i);
System.out.println("is in table?"+para.isInTable());
System.out.println("is in tablelevel?"+para.getTableLevel());
System.out.println("is in isTableRowEnd?"+para.isTableRowEnd());
if(removeEnter(para.text()).contains("样车")){
	
}
// if (para.isInList()) {  
 System.out.println("list: " + para.text());  
//}  
} */
///////////////////////////////////////////////
	      
	      while (tables.hasNext()) {
	    	 Table table = tables.next();
	         //表格属性
//	       CTTblPr pr = table.getCTTbl().getTblPr();
	         //获取表格对应的行
	    	  //获取表格的列数
	    	 int rowsNum = table.numRows();
	         System.out.println("rowsnum=="+rowsNum);
	         for(int i=0;i<rowsNum;i++){
	        	 TableRow row = table.getRow(i);
	        	 int colsNum=row.numCells();
	        	 System.out.println("colsnum="+colsNum);
	         //判断表格
	        	 int ii=0;
	        	 int flag=0;
	        	 if(colsNum%2!=0&&colsNum>3){
	        		 colsNum--;
	        		 ii++;
	        		 flag=1;
	        	 }
	         if(colsNum%2==0){
	        	 
	        	 if(flag==1)
	        		 {
	        		 colsNum++;
	        		 flag=0;
	        		 }
	        	 //for (XWPFTableRow row : rows) {
	 	            //获取行对应的单元格
	 	  
	 	           
	 	         	for(;ii<colsNum-1;ii++){
	 	         		/*String strr=null;
	 	         		for(int hh=0;hh<row.getCell(ii).text().length();hh++){
	 	         			strr+=Integer.toHexString((int)row.getCell(ii).text().charAt(hh));
	 	         			
	 	         		}
	 	         		System.out.println(strr);*/
/*
		                System.out.println("exzt="+matcher("[^[\u4E00-\u9FA5]|/w|,|//]",
		                		row.getCell(ii).text());*/
	 	         		/* System.out.println("exzt="+matcher("\\u0020",
	 	         				stringToUnicode(row.getCell(ii).text())).find());
		                System.out.println(
		                		row.getCell(ii).text().replaceAll("\\u0020", ""));
		                String newStr = new String(row.getCell(ii).text().getBytes()); */
		                //System.out.print(newStr);
		                //System.out.println("fffff="+WordToTextConverter(row.getCell(ii).text()));
		                //System.out.println(unicodeToString(stringToUnicode(row.getCell(ii).text()).replaceAll("$\7", "bb")));
		                //System.out.println(row.getCell(ii).text().substring(0, row.getCell(ii).text().length()-1));
		                
		               /*params.put(matcher("\r\n",row.getCell(ii).text()).replaceAll(""), 
		            		   matcher("\r\n",row.getCell(++ii).text()).replaceAll(""));*/
	 	         		
    	        		
		               if(removeEnter(row.getCell(ii).text())!=null){
		                
		               params.put(removeEnter(row.getCell(ii).text()), 
		            		   removeEnter(row.getCell(++ii).text()));
		               //////////////////////////
		               params.put(removeEnter(row.getCell(0).text())+removeEnter(row.getCell(ii-1).text()), 
		            		   removeEnter(row.getCell(ii).text()));
		               }
		               else{
		            	   ii++;
		               }
		               // System.out.println("gggggggggggg="+cells.get(ii).getText()+"hh"+cells.get(++ii).getText());

		               // System.out.println("params.size0="+params.size());
		               
		            }
	 	         } 
	         
	         else {
	        		System.out.println("colsnum="+colsNum);
	        		boolean lastrow=i+2>rowsNum?true:false;
	        		System.out.println("lstrow="+lastrow);
		        		if(colsNum==1&&lastrow
		        			||colsNum==3&&lastrow
		        			||colsNum==1&&!lastrow&&table.getRow(i+1).numCells()>1
		        			||colsNum==3&&!lastrow&&table.getRow(1+i).numCells()!=3){
			        		System.out.println("read into break");
			        		//if(lastrow)System.out.println("next col num="+table.getRow(1+i).numCells());
			        	
			        	}else{
			        		
			        		
			        		System.out.println("out break");
			        		  for(int j=0;j<colsNum;j++){
			        			  
			        			  //System.out.println("fdf="+matcher("d",table.getRow(1+i).getCell(j).text()).find());
		        		//System.out.println(rows.get(i).getTableCells().get(j).getText()+"           "+rows.get(1+i).getTableCells().get(j).getText());
			      /*params.put(matcher("//s*|/t|/r|/n",table.getRow(i).getCell(j).text()).replaceAll(""),
			    		  matcher("//s*|/t|/r|/n",table.getRow(1+i).getCell(j).text()).replaceAll(""));*/
			        			  
			        			  params.put(removeEnter(table.getRow(i).getCell(j).text()),
			    			    		  removeEnter(table.getRow(1+i).getCell(j).text()));
			      
			        	}
			        		i++;
			        		 
			        	}         
		        			
	
	         }       
		            
	         }
	         
	      } 

	      this.close(is);
			 }
		 }
	   
	      
	   }
	   /**
	    * 关闭输入流
	    * @param is
	    */
	   private void close(InputStream is) {
	      if (is != null) {
	         try {
	            is.close();
	         } catch (IOException e) {
	            e.printStackTrace();
	         }
	      }
	   }
	   
	   
	   private Matcher matcher(String regex,String str) {
		     //Pattern pattern = Pattern.compile("s1023", Pattern.CASE_INSENSITIVE);
		  Pattern pattern = Pattern.compile(regex,Pattern.CASE_INSENSITIVE);
		     Matcher matcher = pattern.matcher(str);
		     return matcher;
		  }
	   public static String stringToUnicode(String s) {
			String str = "";
			for (int i = 0; i < s.length(); i++) {
				int ch = (int) s.charAt(i);
				if (ch > 255)
					str += "\\u" + Integer.toHexString(ch);
				else
	 				str += "\\" + Integer.toHexString(ch);
			}
			return str;
		}
	   public static String unicodeToString(String str) {

	        Pattern pattern = Pattern.compile("(\\\\u(\\p{XDigit}{4}))");    

	        Matcher matcher = pattern.matcher(str);

	        char ch;

	        while (matcher.find()) {

	            ch = (char) Integer.parseInt(matcher.group(2), 16);

	            str = str.replace(matcher.group(1), ch + "");    

	        }

	        return str;

	    }
	   
 public String removeEnter(String str){
	 
	 String tmp=str.substring(0, str.length()-1);
	 if(tmp==null){
		 tmp="";
	 }
	 return tmp;
 }
	}
