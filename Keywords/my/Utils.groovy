package my

import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.filechooser.FileSystemView

import com.kms.katalon.core.annotation.Keyword

public class Utils {
	
	@Keyword
	public static int showFileChooser(){
		int returnValue = 0;
		JFrame frame = new JFrame("title")
		frame.setAlwaysOnTop(true)
		frame.requestFocus()
		File file=new File("./docs")
		JFileChooser j=new JFileChooser(file, FileSystemView.getFileSystemView())
		returnValue=j.showOpenDialog(frame)
		return returnValue
	}
}
