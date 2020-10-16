/*
 * Copyright Public Record Office Victoria 2018
 * Licensed under the CC-BY license http://creativecommons.org/licenses/by/3.0/au/
 * Author Andrew Waugh
 * Version 1.0 February 2018
 */
package eml2v3;

import java.nio.file.Path;
import java.time.Instant;

/**
 * This class represents the result of processing a VEO.
 */
public class VEOResult {

    public String veoId;            // file name of VEO
    public int veoType;             // type of VEO
    static public int V2_VEO = 1;   // VEO was a V2 VEO
    static public int V3_VEO = 2;   // VEO was a V3 VEO
    public Instant timeProcStart;   // time processing started
    public Instant timeProcEnded;   // time processing ended
    public boolean success;         // true if VEO was successfully processed
    public String result;           // text describing the result of processing
    public Path packages;           // the location of the packages generated

    /**
     * Construct a VEOResult
     *
     * @param veoType type of VEO being processed
     */
    public VEOResult(int veoType) {
        success = false;
        this.veoType = veoType;
        this.timeProcStart = Instant.now();
    }

    /**
     * Construct a VEOResult
     * 
     * @param veoId     file name of VEO
     * @param veoType   type of VEO 
     * @param success   true if VEO was successfully processed
     * @param result    test describing result of processing
     * @param packages  location of the DAS, AMS, & SAMS packages (null if result is false)
     * @param started   instant VEO processing started 
     */
    public VEOResult(String veoId, int veoType, boolean success, String result, Path packages, Instant started) {
        this.veoId = veoId;
        this.veoType = veoType;
        this.timeProcStart = started;
        timeProcEnded = Instant.now();
        this.success = success;
        this.result = result;
        this.packages = packages;
    }

    /**
     * Free the VEO Result and all who sail in her
     */
    public void free() {
        veoId = null;
        timeProcStart = null;
        timeProcEnded = null;
        result = null;
        packages = null;
    }
}
