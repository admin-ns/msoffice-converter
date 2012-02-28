To run the server:
    1. Copy the PptConverter.jar file and PptConverter_lib directory to the Windows server that will be hosting the utility.
    2. Be sure port 1099 is available on the server and not blocked by the firewall.
    3. Create a CLASSPATH environment variable and add "C:\<path>\<to>\PptConverter.jar" to it.
    4. AFTER the environment var is created, open a cmd window and run "start rmiregistry" (assuming "\<path_to_jre>\bin\" is in your PATH environment variable).
    5. In the cmd window, run "start java PptConverterServer"


To run the client:
    1. Copy the PptConverter.jar file and PptConverter_lib directory to desired server.
    2. Run "java -jar /<path>/<to>/PptConverter.jar <windows_host> <path_to_ppt_file> <path_to_output_dir>
