

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;

import com.healthmarketscience.rmiio.GZIPRemoteInputStream;
import com.healthmarketscience.rmiio.RemoteInputStreamServer;
import com.healthmarketscience.rmiio.RemoteOutputStreamServer;
import com.healthmarketscience.rmiio.SimpleRemoteOutputStream;


public class PptConverterClient {

	public static void main (String[] args) {
		if (args.length < 3) {
			System.out.println("Usage: java -jar PptConverter.jar <host> <file_to_convert> <output_directory>");
			System.exit(0);
		}
		
		String host = args[0];
		String pptFilePath = args[1];
		String outputFolderPath = args[2];

		RemoteInputStreamServer istream = null;
		RemoteOutputStreamServer ostream = null;
		try {
			Registry registry = LocateRegistry.getRegistry(host);
			PptConverter presentationConverter = (PptConverter) registry.lookup(PptConverterServer.PRESENTATION_CONVERTER_NAME);
			
			istream = createRemoteInputFileStream(pptFilePath);
			
			File outputFolder = new File(outputFolderPath);
			if (!outputFolder.exists()) {
				if (!outputFolder.mkdirs())
					throw new RuntimeException("Unable to create output directory '" + outputFolderPath + "'");
			}
			
			File outputImagesZip = new File(outputFolder, "images.zip");
			ostream = createRemoteOutputFileStream(outputImagesZip);
			
			presentationConverter.convertToImages(istream.export(), ostream.export());

			ZipUtils.unzip(outputImagesZip, outputFolder);
		}
		catch(Exception e){
			e.printStackTrace();
			System.exit(1);
		}
		finally {
			if (istream != null)
				istream.close();
			if (ostream != null)
				ostream.close();
		}
	}

	private static RemoteOutputStreamServer createRemoteOutputFileStream(File outputFile) throws IOException, FileNotFoundException {
		RemoteOutputStreamServer ostream;
		ostream = new SimpleRemoteOutputStream(new BufferedOutputStream(new FileOutputStream(outputFile)));
		return ostream;
	}

	private static RemoteInputStreamServer createRemoteInputFileStream(String pptFilePath) throws IOException, FileNotFoundException {
		RemoteInputStreamServer istream;
		File inputFile = new File(pptFilePath);
		BufferedInputStream bufferedStream = new BufferedInputStream(new FileInputStream(inputFile));
		istream = new GZIPRemoteInputStream(bufferedStream);
		return istream;
	}
}
