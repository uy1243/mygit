package com.yu;


public class Replacedoc {
	ListFile lf=new ListFile();
	/*Writedoc wdoc=new Writedoc();
    wdoc.testWriteTable();*/
	//search pwd docs
	
    
	private  List<String> fls=new ArrayList<String>();


/**
   * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
   * @throws Exception
   */
  public  Replacedoc(Map<String, String> params,List<String> fpath) throws Exception {
	    System.out.println("params="+params.toString());

	  //遍历读入内存，放入Map中
	  
	  //fls=lf.listFile("", "", "");
	  for(String path:fpath){
		  System.out.println(path);
		  
	   if(path.contains("模板")){
  
     //Map<String, Object> params = new HashMap<String, Object>();
     //params.put("s1023", "100.00");
     //加载模板
     //String filePath = "new.docx";
     InputStream is = new FileInputStream(path);
     HWPFDocument doc = new HWPFDocument(is);
     
     //替换内容
     //替换段落里面的变量
    // this.replaceInDoc(doc, params);
     //替换表格里面的变量
     this.replaceInTable(doc, params);
     
     //写入新文件
     File tmpFile=new File(path);
     //System.out.println("filenameeeeeeeeeeeeeeeeee="+matcherss("模板|[0-9]|[a-z]|[A-Z]|-|[/.]",tmpFile.getName()).replaceAll(""));
     Pattern pattern = Pattern.compile("([^\"|[\u4e00-\u9fa5]]+)");
     Matcher matcher = pattern.matcher(tmpFile.getName());
     
     OutputStream os = new FileOutputStream(params.get("车型")+"("+params.get("工号")+"-"+params.get("发动机型号")+")"+
    matcherss("模板|[0-9]|[a-z]|[A-Z]|-|[/.]|[/(]|[/)]",tmpFile.getName()).replaceAll("")+".doc");
     doc.write(os);
     this.close(os);
     this.close(is);
  }
	  }
  }
  /**
   * 替换段落里面的变量
   * @param doc 要替换的文档
   * @param params 参数
   */
  private void replaceInDoc(XWPFDocument doc, Map<String, String> params) {
     Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
     XWPFParagraph para;
     while (iterator.hasNext()) {
        para = iterator.next();
        this.replaceInPara(para, params);
     }
  }
 

  /**
   * 替换表格里面的变量
   * @param doc 要替换的文档
   * @param params 参数
   */
  private void replaceInTable(HWPFDocument doc, Map<String, String> params) {
	
	  Range range = doc.getRange();   
      TableIterator tables = new TableIterator(range);
     while (tables.hasNext()) {
    	 Table table = tables.next();

    	 int rowsNum = table.numRows();

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
    	    	        
    	    	        for(;ii<colsNum-1;ii++){
	               

 	                //System.out.println(cells.get(i).getText());
    	    	        	for (String vv : params.keySet()) {
    	    	        		 
    	    	        		//System.out.println("binary==="+StrToBinstr(row.getCell(ii).text()));
    	    	        		//System.out.println("binary==="+row.getCell(ii).text());
    	    	        		//System.out.println("binary==="+remove111(row.getCell(ii).text()));
	   	            	//System.out.println("vv===="+vv+",,,text="+removeEnter(row.getCell(ii).text()));
    	    	        	if(vv.equals(removeEnter(row.getCell(ii).text()))){
	   	            		 //System.out.println("vv="+vv+",text="+cells.get(++ii).getText());
	   	            		 //row.getCell(++ii).replaceText(params.get(vv),true);
	   	            		//System.out.println("vv===="+row.getCell(ii).text());
	   	            		//System.out.println("vv+1===="+row.getCell(ii+1).text()); 	     
	   	            		//System.out.println("vv+111===="+remove111(row.getCell(ii+1).text())); 
	   	            		//System.out.println("vvpara===="+params.get(vv));
	   	            		//System.out.println("vvpara===="+addEnter(params.get(vv)));
	   	            		if(remove111(row.getCell(ii+1).text())!=""){
	   	            		 
	   	            		 if(removeEnter(row.getCell(ii+1).text())!=params.get(vv)){
	   	            			System.out.println("fffsfsadfsdfsdf");
	   	            			row.getCell(ii+1).replaceText(remove111(row.getCell(ii+1).text()),params.get(vv));
	   	            		 }
	   	            			ii++;
	   	            		 System.out.println("out");
	   	            		}else{
	   	            			ii++;
	   	            		}
	   	            		
	   	            	 	}   
	   	            	 }
    	    	 
	            }
 	         } 
         
/*         else{ 
        	 
        	 
        	 if(colsNum==1&&table.getRow(i+1).numCells()>1
	        			||colsNum==3&&table.getRow(1+i).numCells()!=3){
		        		System.out.println("read into break");
		        	}else{
		        		
		        		System.out.println("out break");
		        		  for(int j=0;j<colsNum;j++){
		        			  for (String vv : params.keySet()) {
		      	            	 //System.out.println("vv===="+vv+",,,text="+cells.get(i).getText());
		      	            	 if(vv.equals(removeEnter(table.getRow(i).getCell(j).text().toString()))){
		      	            		table.getRow(1+i).getCell(j).replaceText("",true);
		         		              table.getRow(++i).getCell(j).replaceText(params.get(vv), true);
		      	            	 }    
		      	           }
               
		        	}
		        	i++;	 
		        	}         
         }  */ 
    	        ///////////////////////////////////////
    	        else {
	        		System.out.println("colsnum="+colsNum);
	        		boolean lastrow=i+2>rowsNum?true:false;
	        		System.out.println("lstrow="+lastrow);
		        		if(colsNum==1&&lastrow
		        			||colsNum==3&&lastrow
		        			||colsNum==1&&!lastrow&&table.getRow(i+1).numCells()>1
		        			||colsNum==3&&!lastrow&&table.getRow(1+i).numCells()!=3){
			        		System.out.println("read into break");
			        		
			        	}else{
			        		
			        		
			        		System.out.println("out break");
			        		  for(int j=0;j<colsNum;j++){
			        			  
			        			  //System.out.println("fdf="+matcher("d",table.getRow(1+i).getCell(j).text()).find());
		        		//System.out.println(rows.get(i).getTableCells().get(j).getText()+"           "+rows.get(1+i).getTableCells().get(j).getText());
			      /*params.put(matcher("//s*|/t|/r|/n",table.getRow(i).getCell(j).text()).replaceAll(""),
			    		  matcher("//s*|/t|/r|/n",table.getRow(1+i).getCell(j).text()).replaceAll(""));*/
			        			  System.out.println("table cccc="+table.getRow(i).getCell(j).text());
			        			  System.out.println("table cccc="+table.getRow(i+1).getCell(j).text());
			        			  System.out.println("j="+j);
			        			  System.out.println("i="+i);
			        			  for (String vv : params.keySet()) {
		    	    	        	if(vv.equals(removeEnter(table.getRow(i).getCell(j).text()))){
		    	    	        		 	            		
			   	            		if(remove111(table.getRow(i+1).getCell(j).text())!=""){
			   	            		row = table.getRow(i+1);
			   	            			table.getRow(i+1).getCell(j).replaceText(remove111(row.getCell(j).text()),params.get(vv));
			   	            			//table.getRow(1+i).getCell(j).replaceText("",true);
			         		              //table.getRow(++i).getCell(j).replaceText(params.get(vv), true);
			   	            		}
			   	            		else{
			   	            			//j++;
			   	            		}
			   	            		
			   	            	 	}   
			   	            	 }
			        			  
			        			  
			      
			        	}
			        		i++;
			        		 
			        	}         
		        			
	
	         } 
    	        /////////////////////////////
	            
         }
         
    			 }
     }
               
 	            
    	            

 
  /**
   * 正则匹配字符串
   * @param str
   * @return
   */
/*  private Matcher matcher(String str) {
     Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
     Matcher matcher = pattern.matcher(str);
     return matcher;
  }*/
  private Matcher matcher(String str) {
	     //Pattern pattern = Pattern.compile("s1023", Pattern.CASE_INSENSITIVE);
	  Pattern pattern = Pattern.compile(str, Pattern.CASE_INSENSITIVE);
	     Matcher matcher = pattern.matcher(str);
	     return matcher;
	  }
  /**
   * 替换段落里面的变量
   * @param para 要替换的段落
   * @param params 参数
   */
  private void replaceInPara(XWPFParagraph para, Map<String, String> params) {
     List<XWPFRun> runs;
     Matcher matcher;
     if (this.matcher(para.getParagraphText()).find()) {
        runs = para.getRuns();
        for (int i=0; i<runs.size(); i++) {
           XWPFRun run = runs.get(i);
           String runText = run.toString();
           matcher = this.matcher(runText);
           if (matcher.find()) {
               while ((matcher = this.matcher(runText)).find()) {
            	   //System.out.print("hello11!");
                  //runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
                  runText = matcher.replaceAll(String.valueOf(params.get(matcher.group(2))));
               }
               //直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
               //所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
               para.removeRun(i);
               para.insertNewRun(i).setText(runText);
           }
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
 
  /**
   * 关闭输出流
   * @param os
   */
  private void close(OutputStream os) {
     if (os != null) {
        try {
           os.close();
        } catch (IOException e) {
           e.printStackTrace();
        }
     }
  }
  public String removeEnter(String str){
		 
		 String tmp=str.substring(0, str.length()-1);
		 return tmp;
	 }
  public String addEnter(String str){
		 
		 String tmp=BinstrToStr(StrToBinstr(str)+" "+"111");
		 //System.out.println("rmp="+tmp);
		 return tmp;

	 }
  
  private Matcher matcherss(String regex,String str) {
	     //Pattern pattern = Pattern.compile("s1023", Pattern.CASE_INSENSITIVE);
	  Pattern pattern = Pattern.compile(regex);
	     Matcher matcher = pattern.matcher(str);
	     return matcher;
	  }
  
  private String StrToBinstr(String str) {
      char[] strChar=str.toCharArray();
      String result="";
      for(int i=0;i<strChar.length;i++){
          result +=Integer.toBinaryString(strChar[i])+ " ";
      }
      return result;
  }
  private String remove111(String str){
	  char[] strChar=str.toCharArray();
	  str="";
	  for(int i=0;i<strChar.length-1;i++){
          str+=strChar[i];
      }
      return str;
  }
  private String BinstrToStr(String binStr) {
      String[] tempStr=StrToStrArray(binStr);
      char[] tempChar=new char[tempStr.length];
      for(int i=0;i<tempStr.length;i++) {
          tempChar[i]=BinstrToChar(tempStr[i]);
      }
      return String.valueOf(tempChar);
  }
  private char BinstrToChar(String binStr){
      int[] temp=BinstrToIntArray(binStr);
      int sum=0;   
      for(int i=0; i<temp.length;i++){
          sum +=temp[temp.length-1-i]<<i;
      }   
      return (char)sum;
  }
  private int[] BinstrToIntArray(String binStr) {       
      char[] temp=binStr.toCharArray();
      int[] result=new int[temp.length];   
      for(int i=0;i<temp.length;i++) {
          result[i]=temp[i]-48;
      }
      return result;
  }
  private String[] StrToStrArray(String str) {
      return str.split(" ");
  }
}
