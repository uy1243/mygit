package com.yu;


import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ListFile {
	private  String filepre;//文件前缀
	private  String filesux;//文件后缀
	private  List<String> fls=new ArrayList<String>();
	public  List<String> listFile(String dir, String prefix, String suffix){
		File fileTarget = new File(dir);//取得目标目录
		if(!fileTarget.isDirectory()){
			//fileTarget=new File("/home/yu/workspace/forw");
			fileTarget=fileTarget.getParentFile();
			//System.out.println(fileTarget.isDirectory()+"  "+fileTarget.exists());
		}
		filepre = prefix;
		filesux =suffix;
		if(fileTarget.exists()){//判断目录是否存在
			/*File[] fileLogs = fileTarget.listFiles(
					new FilenameFilter(){
						public boolean accept(File dir, String name) {
							return (name.endsWith(filesux));//使用FilenameFilter类过滤取得满足指定条件的文件的文件数组
						}	
					}
					);*/
			File[] fileLogs = fileTarget.listFiles();
			for(int i=0;i<fileLogs.length;i++){
				//System.out.println(fileLogs[i].getName());
				if(fileLogs[i].getName().endsWith(filesux)){
					fls.add(fileLogs[i].getAbsolutePath());
					//System.out.println(fls.get(0));
				}
			}
			/*if(fileLogs.length > 0){
				for(int i = 0; i<fileLogs.length; i++){
					System.out.println(fileLogs[i].getName());
				}				
			}else{
				System.err.println("we cant find the file start with:" + prefix);
				System.exit(0);
			}*/
		}else{
			System.err.print("we cant find the path:" + fileTarget);
		    }
		return fls;
		
	}
}


