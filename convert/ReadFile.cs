using System;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
namespace convert
{
    public class ReadFile
    {
        private static String path = "excel_patch";
        private static bool _dataIsAvailableOnDrive;
        private static string _resolvedDataDir;
        public ReadFile()
        {
        }
        public static Stream Exccel(String FileName)
        {
            Initialise();

            if (_dataIsAvailableOnDrive)
            {
                Stream result = OpenClasspathResource(FileName);
                if (result == null)
                {
                    throw new Exception("specified test sample file '" + FileName
                            + "' not found on the classpath");
                }
                //          System.out.println("Opening cp: " + sampleFileName);
                // wrap to avoid temp warning method about auto-closing input stream
                return new NonSeekableStream(result);
            }
            if (_resolvedDataDir == "")
            {
                throw new Exception("Must set system property '"
                                    + path
                        + "' properly before running tests");
            }


            if (!File.Exists(_resolvedDataDir + FileName))
            {
                throw new Exception("Sample file '" + FileName
                        + "' not found in data dir '" + _resolvedDataDir + "'");
            }


            //      System.out.println("Opening " + f.GetAbsolutePath());
            try
            {
                return new FileStream(_resolvedDataDir + FileName, FileMode.Open);
            }
            catch (FileNotFoundException)
            {
                throw;
            }
        }

        private static void Initialise()
        {
            String dataDirName = System.Configuration.ConfigurationSettings.AppSettings[path];

            if (dataDirName == "")
                throw new Exception("Must set system property '"
                                    + path
                        + "' before running tests");

            if (!Directory.Exists(dataDirName))
            {
                throw new IOException("Data dir '" + dataDirName
                        + "' specified by system property '"
                                      + path + "' does not exist");
            }
            _dataIsAvailableOnDrive = true;
            _resolvedDataDir = dataDirName;
        }
        private static Stream OpenClasspathResource(String sampleFileName)
        {
            FileStream file = new FileStream(System.Configuration.ConfigurationSettings.AppSettings["HSSF.testdata.path"] + sampleFileName, FileMode.Open);
            return file;
        }
        \}

        private class NonSeekableStream : Stream
        {

            private Stream _is;

            public NonSeekableStream(Stream is1)
            {
                _is = is1;
            }

            public int Read()
            {
                return _is.ReadByte();
            }
            public override int Read(byte[] b, int off, int len)
            {
                return _is.Read(b, off, len);
            }
            public bool markSupported()
            {
                return false;
            }
            public override void Close()
            {
                _is.Close();
            }
            public override bool CanRead
            {
                get { return _is.CanRead; }
            }
            public override bool CanSeek
            {
                get { return false; }
            }
            public override bool CanWrite
            {
                get { return _is.CanWrite; }
            }
            public override long Length
            {
                get { return _is.Length; }
            }
            public override long Position
            {
                get { return _is.Position; }
                set { _is.Position = value; }
            }
            public override void Write(byte[] buffer, int offset, int count)
            {
                _is.Write(buffer, offset, count);
            }
            public override void Flush()
            {
                _is.Flush();
            }
            public override long Seek(long offset, SeekOrigin origin)
            {
                return _is.Seek(offset, origin);
            }
            public override void SetLength(long value)
            {
                _is.SetLength(value);
            }
        }

        public static HSSFWorkbook OpenSampleWorkbook(String sampleFileName)
        {
            try
            {
                return new HSSFWorkbook(OpenSampleFileStream(sampleFileName));
            }
            catch (IOException)
            {
                throw;
            }
        }
        /**
         * Writes a spReadsheet to a <tt>MemoryStream</tt> and Reads it back
         * from a <tt>ByteArrayStream</tt>.<p/>
         * Useful for verifying that the serialisation round trip
         */
        public static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {

            try
            {
                MemoryStream baos = new MemoryStream(4096);
                original.Write(baos);
                return new HSSFWorkbook(baos);
            }
            catch (IOException)
            {
                throw;
            }
        }

        /**
         * @return byte array of sample file content from file found in standard hssf test data dir 
         */
        public static byte[] GetTestDataFileContent(String fileName)
        {
            MemoryStream bos = new MemoryStream();

            try
            {
                Stream fis = HSSFTestDataSamples.OpenSampleFileStream(fileName);

                byte[] buf = new byte[512];
                while (true)
                {
                    int bytesRead = fis.Read(buf, 0, buf.Length);
                    if (bytesRead < 1)
                    {
                        break;
                    }
                    bos.Write(buf, 0, bytesRead);
                }
                fis.Close();
            }
            catch (IOException)
            {
                throw;
            }
            return bos.ToArray();
        }


    }




}
