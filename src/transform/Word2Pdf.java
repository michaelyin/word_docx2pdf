package transform;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

//https://blog.csdn.net/m0_37568521/article/details/78545887
//https://www.jianshu.com/p/76a8228813c9

public class Word2Pdf {
	public static void main(String args[]) {
		ActiveXComponent app = null;
		String wordFile = "C:\\Users\\michael\\Downloads\\InvoiceTpl.docx";
		String pdfFile = "C:\\Users\\michael\\Downloads\\InvoiceTpl.pdf";
		System.out.println("start ...");
		long start = System.currentTimeMillis();
		try {
			// app = new ActiveXComponent("Word.Application"); // for microsoft office installation
			app = new ActiveXComponent("KWPS.Application");  // for wps installation
			Dispatch documents = app.getProperty("Documents").toDispatch();
			System.out.println("convert: " + wordFile);
			Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
			File target = new File(pdfFile);
			if (target.exists()) {
				target.delete();
			}
			System.out.println("generate: " + pdfFile);
			Dispatch.call(document, "SaveAs", pdfFile, 17);
			Dispatch.call(document, "Close", false);
			long end = System.currentTimeMillis();
			System.out.println("word 2 pdf takes " + (end - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("convert Exception: " + e.getMessage());
		} finally {
			app.invoke("Quit", 0);
		}
	}
}