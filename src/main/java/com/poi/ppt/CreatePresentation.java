package com.poi.ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * 
 * @author Naveen
 *
 */
public class CreatePresentation {

	 public static void main(String ar[]) throws IOException {
		   
	      //opening an existing slide show
		  File file = new File("d://example1.pptx");
		  if(!file.exists()) {
			  //creating a new empty slide show
		      XMLSlideShow ppt = new XMLSlideShow();	   
		      FileOutputStream out = new FileOutputStream(file);
		      ppt.write(out);
		      out.close();
		  }
	      FileInputStream inputstream = new FileInputStream(file);
	      XMLSlideShow ppt = new XMLSlideShow(inputstream);
	      
	      //adding slides to the slodeshow
	      XSLFSlide slide1 = ppt.createSlide();
	      XSLFSlide slide2 = ppt.createSlide();
	      
	      //saving the changes 
	      FileOutputStream out = new FileOutputStream(file);
	      ppt.write(out);
	      
	      System.out.println("Presentation edited successfully");
	      out.close();	
	   }
}
