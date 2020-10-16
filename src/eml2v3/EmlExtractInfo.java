/**
 * ***********************************************************
 *
 * E M L   E X T R A C T  I N F O
 *
 * This class processes emails in EML files and extracts information from them
 *
 * <li><b>[files|directories]+</b> a list of EML files, or directories
 * containing such files.</li>
 * </ul>
 * <p>
 * The following command line arguments are optional:
 * <li><b>-v</b> verbose output. By default off.</li>
 * <li><b>-d</b> debug mode. In this mode more logging will be generated. By
 * default off.</li>
 * <li><b>-o &lt;outputDir&gt;</b> the directory in which extracted information
 * is produced</li>
 * </ul>
 * <p>
 * A minimal example of usage is<br>
 * <pre>
 *     emlextractinfo veo1.veo
 * </pre>
 *
 * Copyright Public Record Office Victoria 2020 Licensed under the CC-BY license
 * http://creativecommons.org/licenses/by/3.0/au/ Author Andrew Waugh Version
 * 1.0 February 2018
 */
package eml2v3;

import VERSCommon.AppFatal;
import VERSCommon.AppError;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;

public class EmlExtractInfo {

    static String classname = "EmlExtractInfo"; // for reporting
    ArrayList<String> files;// files or directories to preProcessEML
    Runtime r;

    // global variables storing information about this export (as a whole)
    Path outputDirectory;   // directory in which VEOS are to be generated
    int emailCount;         // number of exports processed
    boolean debug;          // true if in debug mode
    boolean verbose;        // true if in verbose output mode

    FileWriter outfw;      // file writer for output
    BufferedWriter outbw;

    private final static Logger LOG = Logger.getLogger("EmlExtractInfo.EmlExtractInfo");

    /**
     * Default constructor
     *
     * @param args arguments passed to program
     * @throws AppFatal if a fatal error occurred
     */
    public EmlExtractInfo(String args[]) throws AppFatal {

        // Set up logging
        System.setProperty("java.util.logging.SimpleFormatter.format", "%4$s: %5$s%n");
        LOG.setLevel(Level.INFO);

        // set up default global variables
        files = new ArrayList<>();

        outputDirectory = Paths.get(".");
        emailCount = 0;
        debug = false;
        verbose = false;
        r = Runtime.getRuntime();

        // process command line arguments
        configure(args);

        // open output file
        try {
            outfw = new FileWriter(outputDirectory.resolve("emlOutput.xml").toFile());
            outbw = new BufferedWriter(outfw);
            outbw.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\" ?>");
            outbw.write("<Emails>\n");
        } catch (IOException ioe) {
            throw new AppFatal("Failure opening output file: " + ioe.getMessage());
        }
    }

    /**
     * Finalise...
     */
    public void close() {
        try {
            outbw.write("</Emails>\n");
            outbw.close();
            outfw.close();
        } catch (IOException ioe) {
            // ignore
        }
        // vp.free();
    }

    /**
     * Configure
     *
     * This method gets the options for this run of the manifest generator a the
     * command line. See the comment at the start of this file for the command
     * line arguments.
     *
     * @param args[] the command line arguments
     * @param VEOFatal if a fatal error occurred
     */
    private void configure(String args[]) throws AppFatal {
        int i;
        String usage = "EmlExtractInfo [-v] [-d] [-o <directory>] (files|directories)*";

        // preProcessEML command line arguments
        i = 0;
        try {
            while (i < args.length) {
                switch (args[i]) {

                    // verbose?
                    case "-v":
                        verbose = true;
                        LOG.setLevel(Level.INFO);
                        // rootLog.setLevel(Level.INFO);
                        i++;
                        break;

                    // debug?
                    case "-d":
                        debug = true;
                        LOG.setLevel(Level.FINE);
                        // rootLog.setLevel(Level.FINE);
                        i++;
                        break;

                    // '-o' specifies output directory
                    case "-o":
                        i++;
                        outputDirectory = checkFile("output directory", args[i], true, true);
                        i++;
                        break;

                    default:
                        // if unrecognised arguement, print help string and exit
                        if (args[i].charAt(0) == '-') {
                            throw new AppFatal("Unrecognised argument '" + args[i] + "' Usage: " + usage);
                        }

                        // if doesn't start with '-' assume a file or directory name
                        files.add(args[i]);
                        i++;
                        break;
                }
            }
        } catch (ArrayIndexOutOfBoundsException ae) {
            throw new AppFatal("Missing argument. Usage: " + usage);
        }

        // check to see if at least one file or directory is specified
        if (files.isEmpty()) {
            throw new AppFatal("You must specify at least one file or directory to process");
        }

        // LOG generic things
        if (debug) {
            LOG.log(Level.INFO, "Verbose/Debug mode is selected");
        } else if (verbose) {
            LOG.log(Level.INFO, "Verbose output is selected");
        }
        LOG.log(Level.INFO, "Output directory is ''{0}''", outputDirectory.toString());
    }

    /**
     * Check a file to see that it exists and is of the correct type (regular
     * file or directory). The program terminates if an error is encountered.
     *
     * @param type a String describing the file to be opened
     * @param name the file name to be opened
     * @param isDirectory true if the file is supposed to be a directory
     * @param create if true, create the directory if it doesn't exist
     * @throws AppFatal if the file does not exist, or is of the correct type
     * @return the File opened
     */
    private Path checkFile(String type, String name, boolean isDirectory, boolean create) throws AppFatal {
        Path p;

        p = Paths.get(name);

        if (!Files.exists(p)) {
            if (!create) {
                throw new AppFatal(classname, 6, type + " '" + p.toAbsolutePath().toString() + "' does not exist");
            } else {
                try {
                    Files.createDirectory(p);
                } catch (IOException ioe) {
                    throw new AppFatal(classname, 9, type + " '" + p.toAbsolutePath().toString() + "' does not exist and could not be created: " + ioe.getMessage());
                }
            }
        }
        if (isDirectory && !Files.isDirectory(p)) {
            throw new AppFatal(classname, 7, type + " '" + p.toAbsolutePath().toString() + "' is a file not a directory");
        }
        if (!isDirectory && Files.isDirectory(p)) {
            throw new AppFatal(classname, 8, type + " '" + p.toAbsolutePath().toString() + "' is a directory not a file");
        }
        return p;
    }

    /**
     * Preprocess the list of files or directories passed in on the command
     * line.
     */
    public void processFiles() {
        int i;
        String file;

        // go through the list of files
        for (i = 0; i < files.size(); i++) {
            file = files.get(i);
            if (file == null) {
                continue;
            }
            try {
                processFile(Paths.get(file));
            } catch (InvalidPathException ipe) {
                LOG.log(Level.WARNING, "***Ignoring file ''{0}'' as the file name was invalid: {1}", new Object[]{file, ipe.getMessage()});
            }
        }
    }

    /**
     * Process an individual directory or file. If a directory, recursively
     * process all of the files (or directories) in it.
     *
     * @param f the file or directory to preProcessEML
     * @param first this is the first entry in the directory
     */
    private void processFile(Path f) {
        DirectoryStream<Path> ds;

        // check that file or directory exists
        if (!Files.exists(f)) {
            if (verbose) {
                LOG.log(Level.WARNING, "***File ''{0}'' does not exist", new Object[]{f.normalize().toString()});
            }
            return;
        }

        // if file is a directory, go through directory and test all the files
        if (Files.isDirectory(f)) {
            if (verbose) {
                LOG.log(Level.INFO, "***Processing directory ''{0}''", new Object[]{f.normalize().toString()});
            }
            try {
                ds = Files.newDirectoryStream(f);
                for (Path p : ds) {
                    processFile(p);
                }
                ds.close();
            } catch (IOException e) {
                LOG.log(Level.WARNING, "Failed to process directory ''{0}'': {1}", new Object[]{f.normalize().toString(), e.getMessage()});
            }
            return;
        }

        if (Files.isRegularFile(f)) {
            processEML(f);
        } else {
            LOG.log(Level.INFO, "***Ignoring directory ''{0}''", new Object[]{f.normalize().toString()});
        }
    }

    /**
     * Preprocess an EML file to extract information that we will later build
     * collections of emails from.
     *
     * @param emlFile the file containing the email
     */
    private void processEML(Path emlFile) {
        Email e;
        int i;

        // check parameters
        if (emlFile == null) {
            return;
        }

        // get name of email
        String s = emlFile.getFileName().toString();
        if (!s.toLowerCase().endsWith(".eml")) {
            System.out.println("Ignoring '" + emlFile.toString() + "' as it is not an EML file");
            return;
        }

        // preprocess the emlFile file
        try {
            // preprocess email
            System.out.println("Processing: " + emlFile.toString());
            e = new Email(emlFile);
            outbw.write("<e>\n");
            outbw.write(" <f>");
            outbw.write(emlFile.toString());
            outbw.write("</f>\n");
            outbw.write(" <s>");
            outbw.write(e.subject);
            outbw.write("</s>\n");
            outbw.write(" <dt>");
            outbw.write(e.sentDate.format(DateTimeFormatter.ISO_OFFSET_DATE_TIME));
            outbw.write("</dt>\n");
            if (e.from != null) {
                for (i = 0; i < e.from.length; i++) {
                    outbw.write(" <f>");
                    writeValue(e.from[i]);
                    outbw.write("</f>\n");
                }
            }
            if (e.to != null) {
                for (i = 0; i < e.to.length; i++) {
                    outbw.write(" <t>");
                    writeValue(e.to[i]);
                    outbw.write("</t>\n");
                }
            }
            if (e.cc != null) {
                for (i = 0; i < e.cc.length; i++) {
                    outbw.write(" <c>");
                    writeValue(e.cc[i]);
                    outbw.write("</c>\n");
                }
            }
            if (e.bcc != null) {
                for (i = 0; i < e.bcc.length; i++) {
                    outbw.write(" <b>");
                    writeValue(e.bcc[i]);
                    outbw.write("</b>\n");
                }
            }
            if (e.mesgID != null) {
                outbw.write(" <i>");
                writeValue(e.mesgID);
                outbw.write("</i>\n");
            }
            if (e.references != null) {
                for (i = 0; i < e.references.length; i++) {
                    outbw.write(" <rs>");
                    writeValue(e.references[i]);
                    outbw.write("</rs>\n");
                }
            }
            if (e.inReplyTo != null) {
                outbw.write(" <irt>");
                writeValue(e.inReplyTo);
                outbw.write("</irt>\n");
            }
            if (e.threadIndex != null) {
                outbw.write(" <ti>");
                writeValue(e.threadIndex);
                outbw.write("</ti>\n");
            }
            outbw.write("</e>\n");
            // System.out.println(abrvMesgId(e.mesgID)+"->"+abrvMesgId(e.inReplyTo));
            // System.out.println(e.references==null?null:e.references[0]);

            // add to index of Message Ids
            emailCount++;
        } catch (AppError err) {
            LOG.log(Level.WARNING, "Processing VEO ''{0}'' failed because:\n{1}", new Object[]{emlFile.toString(), err.getMessage()});
        } catch (AppFatal err) {
            LOG.log(Level.SEVERE, "System error:\n{0}", new Object[]{err.getMessage()});
        } catch (IOException ioe) {
            LOG.log(Level.SEVERE, "IO error:\n{0}", new Object[]{ioe.getMessage()});
        } finally {
            System.gc();
        }
    }

    /**
     * Low level routine to encode an XML value to UTF-8 and write to the XML
     * document. The special characters ampersand, less than, greater than,
     * single quote and double quote are quoted.
     *
     * @param s string to write to XML document
     * @throws IOException if a fatal error occurred
     */
    public void writeValue(String s) throws IOException {
        String module = "write";
        StringBuilder sb = new StringBuilder();
        int i;
        char c;

        // sanity check
        if (s == null || s.length() == 0) {
            return;
        }

        // quote the special characters in the string
        for (i = 0; i < s.length(); i++) {
            c = s.charAt(i);
            switch (c) {
                case '&':
                    sb.append("&amp;");
                    break;
                case '<':
                    sb.append("&lt;");
                    break;
                case '>':
                    sb.append("&gt;");
                    break;
                case '"':
                    sb.append("&quot;");
                    break;
                case '\'':
                    sb.append("&apos;");
                    break;
                default:
                    sb.append(c);
                    break;
            }
        }
        outbw.write(sb.toString());
    }

/**
 * Main program
 *
 * @param args command line arguments
 */
public static void main(String args[]) {
        EmlExtractInfo eei;

        try {
            eei = new EmlExtractInfo(args);
            eei.processFiles();
            eei.close();
            // tp.stressTest(1000);
        } catch (AppFatal e) {
            System.out.println("Fatal error: " + e.getMessage());
            System.exit(-1);
        }
    }
}
