package com.why.project.poilib;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Created by HaiyuKing
 * Used poi工具类封装
 * 在使用POI写word doc文件的时候我们必须要先有一个doc文件才行，因为我们在写doc文件的时候是通过HWPFDocument来写的，
 * 而HWPFDocument是要依附于一个doc文件的。所以通常的做法是我们先在硬盘上准备好一个内容空白的doc文件，然后建立一个基于该空白文件的HWPFDocument。
 * 之后我们就可以往HWPFDocument里面新增内容了，然后再把它写入到另外一个doc文件中，这样就相当于我们使用POI生成了word doc文件。
 */
public class PoiUtils {

	/**
	 * 生成一个doc文件
	 * @param templetDocPath  模板文件的完整路径
	 * @param targetDocPath 生成的目标文件的完整路径
	 * @param dataMap 替换的数据*/
	public static void writeToDoc(String templetDocPath, String targetDocPath, Map<String,String> dataMap){
		try
		{
			//得到模板doc文件的HWPFDocument对象
			InputStream in = new FileInputStream(templetDocPath);
			writeToDoc(in,targetDocPath,dataMap);
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * 生成一个doc文件，主要用于直接读取asset目录下的模板文件，不用先复制到sd卡中
	 * @param templetDocInStream  模板文件的InputStream
	 * @param targetDocPath 生成的目标文件的完整路径
	 * @param dataMap 替换的数据*/
	public static void writeToDoc(InputStream templetDocInStream, String targetDocPath, Map<String,String> dataMap){
		try
		{
			//得到模板doc文件的HWPFDocument对象
			HWPFDocument HDoc = new HWPFDocument(templetDocInStream);
			// 获取word文本内容，整个文本
			Range range = HDoc.getRange();
			// 替换文本内容，将自定义的$xxx$替换成实际文本
			for(Map.Entry<String, String> entry : dataMap.entrySet())
			{
				range.replaceText(entry.getKey(), entry.getValue());
			}
			//写到另一个文件中
			FileOutputStream out = new FileOutputStream(targetDocPath, true);
			//把doc输出到输出流中
			HDoc.write(out);
			out.close();
			templetDocInStream.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
