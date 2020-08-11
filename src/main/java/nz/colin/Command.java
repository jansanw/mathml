package nz.colin;

import nu.xom.Document;
import nz.colin.mathml.Converter;
import nz.colin.mtef.XMLSerialize;
import nz.colin.mtef.exceptions.ParseException;
import nz.colin.mtef.parsers.MTEFParser;
import nz.colin.mtef.records.MTEF;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;

public class Command {

    /**
     * command entrance
     *
     * @param args
     * @throws IOException
     * @throws ParseException
     */
    public static void main(String[] args) throws IOException, ParseException {
        String operate = args.length > 0 ? args[0] : "null";
        switch (operate) {
            case "-f": {
                String mml = mathtype2MML(args[1]);
                System.out.println(mml);
                break;
            }
            case "-d": {
                File file = new File(args[1]);
                File[] listFiles = file.listFiles();
                for (File subFile : listFiles) {
                    String filename = subFile.toString();
                    if (!subFile.isDirectory() && filename.endsWith(".bin")) {
                        try {
                            String mml = mathtype2MML(filename);
                            System.out.println(filename + "\n" + mml);
                        } catch (Exception e) {
                            // skip error
                            // System.out.println(e);
                        }
                    }
                }

                break;
            }
            default: {
                System.out.println("try: java -jar mathml.jar -f /oleObject.bin or java -jar mathml.jar -d /dir");
            }
        }
    }

    /**
     * Mathtype => MathML
     *
     * @param filename
     * @return
     * @throws IOException
     * @throws ParseException
     */
    private static String mathtype2MML(String filename) throws IOException, ParseException {
        File file = new File(filename);
        InputStream is = new FileInputStream(file);
        POIFSFileSystem poifs = new POIFSFileSystem(is);
        PushbackInputStream pis = new PushbackInputStream(poifs.createDocumentInputStream("Equation Native"));

        pis.read(new byte[28]);

        MTEF mtef = new MTEFParser().parse(pis);
        XMLSerialize serializer = new XMLSerialize();

        mtef.accept(serializer);

        Converter c = new Converter();
        Document mathml = c.doConvert(serializer.getRoot());
        is.close();
        return mathml.toXML().replace("&amp;", "&");
    }
}
