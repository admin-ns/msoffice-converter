

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;
import java.rmi.server.UnicastRemoteObject;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;

import com.healthmarketscience.rmiio.RemoteInputStream;
import com.healthmarketscience.rmiio.RemoteInputStreamClient;
import com.healthmarketscience.rmiio.RemoteOutputStream;
import com.healthmarketscience.rmiio.RemoteOutputStreamClient;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class PptConverterServer implements PptConverter {

	public static final String PRESENTATION_CONVERTER_NAME = "PresentationConverter";
	public static final int DEFAULT_PORT = 1099; // Firewall needs to allow whatever port you end up running this on.

	public static void main(String[] args){
		if (args.length == 0) {
			// Just register with the RMI server
			try {
				PptConverterServer server = new PptConverterServer();
				PptConverter stub = (PptConverter) UnicastRemoteObject.exportObject(server, 0);
				
				Registry registry = LocateRegistry.getRegistry();
				try {
					registry.unbind(PRESENTATION_CONVERTER_NAME);
				}catch(Exception e){}
				registry.bind(PRESENTATION_CONVERTER_NAME, stub);
				System.out.println("Server ready");
			}
			catch (Exception e) {
				System.err.println("Unable to start server");
				e.printStackTrace();
				System.exit(1);
			}
		}
		else {
			// Assume they are running this directly.
			if (args.length < 2) {
				System.out.println("Usage: java -jar PresentationConverter.jar <file_to_convert> <output_directory>");
				System.exit(0);
			}
			String fileName = args[0];
			String outputFileName = args[1];
			new PptConverterServer().convertToImages(fileName, outputFileName);
		}
	}

	@Override
	public void convertToImages(RemoteInputStream pptFileStream, RemoteOutputStream outputFolderStream) {
		File localPptFile = convertToLocalFile(pptFileStream);
		File outputFolder = null;
		try {
			if (localPptFile != null) {
				outputFolder = getImageOutputFolderForFile(localPptFile);
				convertToImages(localPptFile.getAbsolutePath(), outputFolder.getAbsolutePath());
				zipAndCopyLocalFilesToRemoteOutputStream(outputFolder, outputFolderStream);
			}
		}
		finally {
			FileUtils.deleteQuietly(outputFolder);
			FileUtils.deleteQuietly(localPptFile);
		}
	}

	private void zipAndCopyLocalFilesToRemoteOutputStream(File outputFolder, RemoteOutputStream remoteOutputStream) {
		OutputStream outputStream = null;
		FileInputStream zipInputStream = null;
		File outputZipFile = null;
		try {
			outputZipFile = new File(FileUtils.getTempDirectory(), outputFolder.getName() + ".zip");
			ZipUtils.zipDirectory(outputFolder, outputZipFile);
			zipInputStream = new FileInputStream(outputZipFile);
			outputStream = RemoteOutputStreamClient.wrap(remoteOutputStream);
			IOUtils.copy(zipInputStream, outputStream);
		} 
		catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		} 
		catch (IOException e) {
			throw new RuntimeException(e);
		}
		finally {
			if (outputStream != null){
				try {
					outputStream.flush();
					outputStream.close();
				}
				catch (IOException e) {
					e.printStackTrace();
				}
			}
			IOUtils.closeQuietly(zipInputStream);
			FileUtils.deleteQuietly(outputZipFile);
		}
	}

	public void convertToImages(String localPptFilePath, String imageOutputFolderPath) {
		Dispatch presentation = null;
		try {
			ActiveXComponent ppt = new ActiveXComponent("PowerPoint.Application");
			Dispatch presentations = ppt.getProperty("Presentations").toDispatch();
			presentation = Dispatch.call(presentations, "Open", localPptFilePath).toDispatch();
			//Dispatch images = Dispatch.call(presentation, "Export", "ppt", outputFileName).toDispatch();
			Dispatch.call(presentation, "SaveAs", imageOutputFolderPath, new Variant(JPEG_FILE_TYPE));
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		finally {
			if (presentation != null)
				Dispatch.call(presentation, "Close");
		}
	}
	
	/**
	 * Pulls the remote file into a temporary file and then returns that file.
	 * @param inFile
	 * @return
	 */
	private File convertToLocalFile(RemoteInputStream inFile) {
		InputStream istream = null;
		OutputStream outputFileStream = null;
		try {
			istream = RemoteInputStreamClient.wrap(inFile);
			File outputFile = File.createTempFile(""+System.currentTimeMillis(), ".ppt", FileUtils.getTempDirectory());
			outputFileStream = new FileOutputStream(outputFile);
			IOUtils.copyLarge(istream, outputFileStream);
			return outputFile;
		}
	    catch (IOException e) {
			e.printStackTrace();
			return null;
		}
	    finally {
	    	try {
	    		if (outputFileStream != null) {
	    			outputFileStream.flush();
	    			outputFileStream.close();
	    			outputFileStream = null;
	    		}
			}
	    	catch (Exception e) {
				System.err.println("Unable to close output stream due to the following error:");
				e.printStackTrace();
			}
	    	
//	    	try {
//				inFile.close(true);
//			}
//	    	catch (Exception e) {
//				System.err.println("Unable to close remote input stream due to the following error:");
//				e.printStackTrace();
//			}
	    }
    }

	private File getImageOutputFolderForFile(File localPptFile) {
		File outputFolder = new File(localPptFile.getParentFile(), FilenameUtils.getBaseName(localPptFile.getName()));
		if (!outputFolder.mkdirs())
			throw new RuntimeException("Unable to make output directory '" + outputFolder + "'");
		return outputFolder;
	}
}
