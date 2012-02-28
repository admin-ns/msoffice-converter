

import java.rmi.Remote;
import java.rmi.RemoteException;

import com.healthmarketscience.rmiio.RemoteInputStream;
import com.healthmarketscience.rmiio.RemoteOutputStream;

public interface PptConverter extends Remote {

	public static final int JPEG_FILE_TYPE = 17;

	public void convertToImages(RemoteInputStream pptFileStream, RemoteOutputStream outputFolder) throws RemoteException;

}