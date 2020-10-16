/*
 * This class encapsulates a single email. It allows for different email
 * processing code.
 */
package eml2v3;

import VERSCommon.AppError;
import VERSCommon.AppFatal;
import VERSCommon.VEOError;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.mail.Address;
import javax.mail.Header;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimePart;
import javax.mail.internet.MimeUtility;

/**
 *
 * @author Andrew
 */
public class Email {

    private int uniqueId;           // counter to give unique file names to parts in temporary directory

    public Path email;             // path of the EML file
    public String recordName;       // name of the record (the EML file name without the .eml)

    boolean dummy;          // true if this email is a dummy (referenced in another email, but not seen

    // extracted information from the email
    public String subject;          // subject of email
    public String[] from;           // from accounts
    public String[] to;             // to accounts
    public String[] cc;             // cc acounts
    public String[] bcc;            // bcc accounts
    public ZonedDateTime sentDate;  // date email sent
    public String mesgID;           // unique email message ID
    public String[] references;     // what email does this one reference?
    public String inReplyTo;        // what email is this one in reply to?
    public String threadIndex;      // Microsoft specific header
    public int hdrCount;            // number of headers in email
    public int lineCount;           // number of lines in content

    // processed contextual information relating to this email
    public int threadLength;        // length of thread
    public ArrayList<Email> replies;// list of replies to this email (threading)
    public Email replyingTo;        // email being replied to
    public boolean brokenThread;    // true if this email replies to an email we know nothing about
    public ArrayList<Email> referencedBy; // list of referencing email

    /**
     * Produce a text representation of this email
     * @return the string
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        int i;

        sb.append(" Email: " + email.toString() + "\n");
        sb.append("  Subject: " + subject + "\n");
        sb.append("  From: ");
        for (i = 0; i < from.length; i++) {
            if (i != 0) {
                sb.append(", ");
            }
            sb.append(from[i]);
        }
        sb.append("\n");
        sb.append("  To: ");
        if (to != null) {
            for (i = 0; i < to.length; i++) {
                if (i != 0) {
                    sb.append(", ");
                }
                sb.append(to[i]);
            }
            sb.append("\n");
        }
        if (cc != null) {
            sb.append("  Cc: ");
            for (i = 0; i < cc.length; i++) {
                if (i != 0) {
                    sb.append(", ");
                }
                sb.append(cc[i]);
            }
            sb.append("\n");
        }
        if (bcc != null) {
            sb.append("  Bcc: ");
            for (i = 0; i < bcc.length; i++) {
                if (i != 0) {
                    sb.append(", ");
                }
                sb.append(bcc[i]);
            }
            sb.append("\n");
        }
        return sb.toString();
    }

    /**
     * Default constructor
     *
     * @param email the path of the EML file
     * @throws VERSCommon.AppFatal an error occurred that means the program has
     * to stop
     * @throws VERSCommon.AppError an error occurred that means processing this
     * email has to stop
     */
    public Email(Path email) throws AppFatal, AppError {
        String s;
        int i;
        Properties p;
        File f;

        free();
        uniqueId = 1;
        threadLength = 0;
        replies = new ArrayList<>();
        referencedBy = new ArrayList<>();

        // get name of email
        s = email.getFileName().toString();
        if (!s.toLowerCase().endsWith(".eml")) {
            throw new AppError("Ignoring '" + email.toString() + "' as it is not an EML file");
        }

        // make sure email path is real...              
        try {
            f = email.toFile().getCanonicalFile();
        } catch (IOException ioe) {
            throw new AppError("Panic! EML file '" + email.toString() + "' was not found: " + ioe.getMessage() + " (Email.constructor)");
        }
        this.email = f.toPath();

        // convert it into the record name
        if ((i = s.lastIndexOf('.')) != -1) {
            recordName = s.substring(0, i);
        } else {
            recordName = s;
        }
        // recordName = recordName.trim() + ".veo.zip";
        recordName = recordName.trim();

        // extract the details of this email
        extractHeaders();
    }

    /**
     * Make a dummy email. This represents an email that is in a thread, but is
     * not part of the collection.
     *
     * @param mesgID the message ID header of the lost email
     * @param recordName the record name from the referencing email
     */
    public Email(String mesgID, String recordName) {
        free();
        uniqueId = 1;
        replies = new ArrayList<>();
        referencedBy = new ArrayList<>();
        dummy = true;
        this.recordName = recordName;
        this.mesgID = mesgID;
    }

    /**
     * Free the email structure and all who sail in her
     */
    final public void free() {
        dummy = false;
        recordName = null;
        subject = null;
        from = null;
        to = null;
        cc = null;
        bcc = null;
        sentDate = null;
        mesgID = null;
        references = null;
        inReplyTo = null;
        threadIndex = null;

        replyingTo = null;
        replies = null;
        referencedBy = null;
    }

    /**
     * Read the email and extract some header information.
     *
     * @throws AppError
     */
    private void extractHeaders() throws AppError {
        MimeMessage m;
        String sa[];
        int i;
        Enumeration<Header> headers;        // all headers

        // parse the email
        m = parseEmail();

        // extract information from email
        try {
            subject = m.getSubject();
            sa = m.getHeader("Date");
            if (sa == null) {
                throw new AppError("Failed in extracting information from '" + email.toString() + "': Email didn't have a Date header (Email.extractHeaders())");
            }
            if (sa.length > 0) {
                // Hack. nsf2x occasionally produces dates with a trailing timezone
                // ones seen include '(GST)', '(UTC)', and '(CST)'. This strips them
                // off
                if ((i = sa[0].lastIndexOf("(")) != -1) {
                    sa[0] = sa[0].substring(0, i-1);
                }
                sentDate = ZonedDateTime.parse(sa[0], DateTimeFormatter.RFC_1123_DATE_TIME);
            } else {
                throw new AppError("Failed in extracting information from '" + email.toString() + "': either 0 or multiple sent dates (Email.extractHeaders())");
            }
            from = addr2Str(m.getFrom());
            to = getRecipients(m, Message.RecipientType.TO);
            cc = getRecipients(m, Message.RecipientType.CC);
            bcc = getRecipients(m, Message.RecipientType.BCC);
            mesgID = m.getMessageID();
            if (mesgID == null) {
                mesgID = getFirstHeader(m, "$MessageID");
            }
            references = getHeader(m, "References");
            inReplyTo = getFirstHeader(m, "In_Reply_To");
            threadIndex = getFirstHeader(m, "Thread_Index");
            if (threadIndex == null) {  // NSF2X uses Thread-Index instead of Thread_Index
                threadIndex = getFirstHeader(m, "Thread-Index");
            }
            headers = m.getAllHeaders();
            i = 0;
            while (headers.hasMoreElements()) {
                i++;
                headers.nextElement();
            }
            hdrCount = i;
            lineCount = m.getLineCount();
        } catch (MessagingException | DateTimeParseException me) {
            throw new AppError("Failed in extracting information from '" + email.toString() + "':" + me.getMessage() + " (Email.extractHeaders())");
        } finally {
            m = null;
        }
    }

    /**
     * Get the from address(es) of the email. Note a from header can contain
     * multiple email addresses
     *
     * @param m parse email
     * @param type type of recepient to return (TO, CC, BCC)
     * @return array of strings containing addresses
     * @throws VERSCommon.AppError something went wrong
     */
    private String[] getRecipients(MimeMessage m, Message.RecipientType type) throws AppError {
        Address[] a;

        // sanity check
        if (type != Message.RecipientType.TO && type != Message.RecipientType.CC && type != Message.RecipientType.BCC) {
            throw new AppError("Invalid message recipient type (Email.getRecipients())");
        }

        // get the sender as an address
        try {
            a = m.getRecipients(type);
        } catch (MessagingException me) {
            a = hackGetRecipients(m, type);
            // throw new AppError("Failed in getting recipient (" + type + ") '" + email.toString() + "':" + me.getMessage() + " (Email.getRecipients())");
        }
        return addr2Str(a);
    }

    /**
     * Hack version of getRecipients(). This is necessary because the EML files
     * produced by nsf2x have a bug: address header lines are folded
     * incorrectly. Specifically, if the address line is greater than some
     * number of characters (998 characters?) a fold ("\r\n ") is brutally
     * inserted irrespective of the address tokens. When the line is unfolded,
     * this means that an extraneous space can appear in an illegal place
     * causing parsing of the address to fail.
     *
     * We 'solve' this equally brutally. We look for the fold ("\r\n ") and
     * simply remove it. This is incorrect according to RFC5322, as you might
     * lose a space.
     *
     * @param m the mime message
     * @param type the type of recipients to get (To, CC, BCC)
     * @return an array of internet addresses
     * @throws AppError if something failed
     */
    private Address[] hackGetRecipients(MimeMessage m, Message.RecipientType type) throws AppError {
        int i;
        Address[] a;
        String sa[];
        String s;

        // get the headers
        sa = null;
        a = null;
        try {
            if (type == Message.RecipientType.TO) {
                sa = m.getHeader("To");
            } else if (type == Message.RecipientType.CC) {
                sa = m.getHeader("Cc");
            } else if (type == Message.RecipientType.BCC) {
                sa = m.getHeader("Bcc");
            } else {
                throw new AppError("Failed in getting recipient list (" + type + ") '" + email.toString() + "': type not To, CC, or Bcc (Email.hackGetRecipients())");
            }
            if (sa == null) {
                return null;
            }

            for (i = 0; i < sa.length; i++) {
                s = sa[i].replace("\r\n ", "");
                s = s.replace("\r\n\t", " ");
                s = s.replace("<'", "<");
                s = s.replace("'>", ">");
                try {
                    a = InternetAddress.parse(s, true);
                } catch (AddressException ae) {
                    throw new AppError("Failed in getting recipient (" + type + ") '" + s + "':" + ae.getMessage() + " (Email.hackGetRecipients())");
                }
            }
        } catch (MessagingException me) {
            throw new AppError("Failed in getting recipient (" + type + ") '" + email.toString() + "':" + me.getMessage() + " (Email.hackGetRecipients())");
        }
        return a;
    }

    /**
     * Decode an array of addresses into an array of Strings. Returns NULL if
     * passed null
     */
    private String[] addr2Str(Address[] a) {
        String r[];
        int i;
        InternetAddress ia;

        if (a == null) {
            return null;
        }
        r = new String[a.length];
        for (i = 0; i < a.length; i++) {
            if (a[i] == null) {
                r[i] = null;
            } else {
                ia = (InternetAddress) a[i];
                try {
                    r[i] = MimeUtility.decodeText(ia.getAddress());
                } catch (UnsupportedEncodingException e) {
                    r[i] = ia.getAddress();
                }
            }
        }
        return r;
    }

    /**
     * Get the first value of an arbitrary email header. If the header does not
     * exist, or has no values, null is returned.
     *
     * @param name the name of the header
     * @return the first value associated with this header
     * @throws VERSCommon.AppError any failures in processing this email
     */
    private String getFirstHeader(MimeMessage m, String name) throws MessagingException {
        String[] headers;

        headers = getHeader(m, name);
        if (headers == null || headers.length < 1) {
            return null;
        }
        return headers[0];
    }

    /**
     * Get an arbitrary email header. Returns NULL if doesn't exist.
     *
     * @param name the name of the header
     * @return an array of Strings containing the values of this header
     * @throws VERSCommon.AppError any failure in processing this email
     */
    private String[] getHeader(MimeMessage m, String name) throws MessagingException {
        String[] headers;
        String s;
        int i;

        // get array of headers matching name
        headers = m.getHeader(name);

        if (headers == null) {
            return null;
        }

        // decode headers if required
        for (i = 0; i < headers.length; i++) {
            try {
                if ((s = headers[i]) != null) {
                    s = MimeUtility.decodeText(s);
                    headers[i] = s;
                }
            } catch (UnsupportedEncodingException e) {
                // ignore - just leave the header encoded as, even encoded,
                // it's still ASCII
            }
        }

        return headers;
    }

    /**
     * Output a String part to a file for inclusion in the VEO.
     *
     * @param tmpDir temporary directory to hold file (will be removed when VEO
     * finalised)
     * @param contents string containing parts
     * @return
     * @throws AppError
     */
    private Path createTempFile(Path tmpDir, String contents, String fileExt) throws AppError {
        Path p;
        FileWriter fw;
        BufferedWriter bw;

        // create file in temporary directory
        p = tmpDir.resolve(Integer.toString(uniqueId) + "." + fileExt);
        if (Files.exists(p)) {
            System.out.println("ARRGH: this file already exists: " + p.toString());
        }
        uniqueId++;
        try {
            fw = new FileWriter(p.toFile());
            bw = new BufferedWriter(fw);
            bw.append(contents);
            bw.close();
            fw.close();
        } catch (IOException ioe) {
            throw new AppError("Failed creating temporary file in breaking down email: " + ioe.getMessage() + "(Email.createTempFile())");
        }
        return p;
    }

    /**
     * Output a non-string part to a file for inclusion in the VEO.
     *
     * @param tmpDir temporary directory to hold file (will be removed when VEO
     * finalised)
     * @param contents string containing parts
     * @return
     * @throws AppError
     */
    private Path createTempFile(Path tmpDir, DataHandler dh, String contentType, String encoding) throws AppError {
        Path p;
        InputStream is;
        BufferedInputStream bis;
        FileOutputStream fos;
        BufferedOutputStream bos;
        byte[] buffer = new byte[1000];
        int i, j;
        String[] tokens;
        String s, name, mimeType;

        // see if there is a 'name' attribute in the contentType
        name = null;
        mimeType = null;
        tokens = contentType.split(";");
        for (i = 0; i < tokens.length; i++) {
            s = tokens[i].trim();
            if (i == 0) {
                mimeType = s;
            }
            if (s.toLowerCase().startsWith("name=")) {
                if ((j = s.indexOf("\"")) == -1) {
                    name = s.substring(5);
                } else {
                    name = s.substring(j + 1, s.length() - 1);
                }
                break;
            }
        }

        // create file in temporary directory
        if (name != null) {
            p = tmpDir.resolve(Integer.toString(uniqueId) + "." + name);
        } else {
            p = tmpDir.resolve(Integer.toString(uniqueId) + "." + "txt");
        }
        uniqueId++;
        try {
            is = dh.getInputStream();
            bis = new BufferedInputStream(is);
            fos = new FileOutputStream(p.toFile());
            bos = new BufferedOutputStream(fos);
            while ((i = is.read(buffer)) != -1) {
                bos.write(buffer, 0, i);
            }
            bos.close();
            fos.close();
            bis.close();
            is.close();
        } catch (IOException ioe) {
            throw new AppError("Failed creating temporary file in breaking down email: " + ioe.getMessage() + "(Email.createTempFile())");
        }
        return p;
    }

    /**
     * Parse email. Note we parse the email several times to reduce the amount
     * of information held in-core to that necessary to do the immediate job.
     * This trades off time vs space and makes it more likely that a large
     * number of emails can be handled.
     */
    private MimeMessage parseEmail() throws AppError {
        Properties p;
        Session session;
        InputStream is;
        BufferedInputStream bis;
        MimeMessage m;          // decoded representation of the email

        // create properties for reading EML file
        p = System.getProperties();
        p.put("mail.host", "smtp.dummydomain.com");
        p.put("mail.transport.protocol", "smtp");

        // get session
        session = Session.getDefaultInstance(p, null);

        // open EML file
        try {
            is = new FileInputStream(email.toFile());
        } catch (FileNotFoundException fnfe) {
            throw new AppError("Panic! EML file '" + email.toString() + "' was not found: " + fnfe.getMessage() + " (Email.constructor)");
        }
        bis = new BufferedInputStream(is);

        // parse EML file
        try {
            m = new MimeMessage(session, bis);
        } catch (MessagingException me) {
            throw new AppError("Failed in parsing EML file '" + email.toString() + "':" + me.getMessage() + " (Email.constructor)");
        } finally {
            try {
                bis.close();
                is.close();
            } catch (IOException ioe) {
                /* ignore */
            }
        }
        return m;
    }

    /**
     * Convert an array of strings into a single string with the original
     * strings separated by ','.
     *
     * @param sa array of Strings
     * @return String
     */
    private String md2Str(String[] sa) {
        StringBuffer sb;
        int i;

        if (sa == null) {
            return "None";
        }
        sb = new StringBuffer();
        for (i = 0; i < sa.length; i++) {
            if (sa[i] != null) {
                if (i != 0) {
                    sb.append(", ");
                }
                sb.append(sa[i]);
            }
        }
        return sb.toString();
    }

    /**
     * Convert an enumeration of headers into XML.
     *
     * @return a StringBuilder containing the XML
     */
    private StringBuilder convHeadersToXML(Enumeration<Header> headers) {
        StringBuilder sb, name;
        Header h;
        int i;
        String s;
        char c;

        sb = new StringBuilder();
        name = new StringBuilder();
        sb.append("<emailHeaders>\n");
        while (headers.hasMoreElements()) {
            h = headers.nextElement();

            // get the header and get rid of any characters that would be
            // illegal in an XML element name
            s = h.getName();
            for (i = 0; i < s.length(); i++) {
                c = s.charAt(i);
                if (Character.isLetter(c)) {
                    name.append(c);
                } else if (Character.isDigit(c)) {
                    if (i == 0) {
                        name.append('X'); //XML element names cannot start with a digit
                    }
                    name.append(c);
                } else if (c == '$') {
                    name.append("Dollar"); // XML element names cannot have '$' characters
                }
            }

            // start the XML element using the header name
            sb.append(" <");
            sb.append(name.toString());
            sb.append(">");

            // output the XML value (header value), converting any special
            // XML characters
            s = h.getValue();
            for (i = 0; i < s.length(); i++) {
                c = s.charAt(i);
                if (c == '&') {
                    sb.append("&amp;");
                } else if (c == '<') {
                    sb.append("&lt;");
                } else if (c == '>') {
                    sb.append("&gt;");
                } else if (c == '"') {
                    sb.append("&quot;");
                } else if (c == '\'') {
                    sb.append("&apos;");
                } else {
                    sb.append(c);
                }
            }

            // end XML element
            sb.append("</");
            sb.append(name.toString());
            sb.append(">\n");

            // reset name StringBuilder for next header
            name.setLength(0);
        }
        sb.append("</emailHeaders>\n");
        return sb;
    }
}
