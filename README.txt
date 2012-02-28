To run the server:
    1. Copy the PptConverter.jar file and PptConverter_lib directory to the Windows server that will be hosting the utility.
    2. Create a CLASSPATH environment variable and add "C:\<path>\<to>\PptConverter.jar" to it.
    3. Open a cmd window and run "start rmiregistry" (assuming "\<path_to_jre>\bin\" is in your PATH environment variable).
    4. In the cmd window, run "start java PptConverterServer"


To run the client:
    1. Copy the PptConverter.jar file and PptConverter_lib directory to desired server.
    2. Run "java -jar /<path>/<to>/PptConverter.jar <windows_host> <path_to_ppt_file> <path_to_output_dir>
