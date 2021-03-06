/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.hmef.attribute;

import org.apache.poi.hmef.Attachment;
import org.apache.poi.hmef.HMEFMessage;
import org.apache.poi.hsmf.datatypes.MAPIProperty;
import org.apache.poi.hsmf.datatypes.Types;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.util.StringUtil;

import java.io.UnsupportedEncodingException;

/**
 * A pure-MAPI attribute holding a String, which applies
 * to a {@link HMEFMessage} or one of its {@link Attachment}s.
 */
public final class MAPIStringAttribute extends MAPIAttribute {
    private static final String CODEPAGE = "CP1252";
    private static POILogger logger = POILogFactory.getLogger(MAPIStringAttribute.class);
    private final String data;

    public MAPIStringAttribute(MAPIProperty property, int type, byte[] data) {
        super(property, type, data);

        String tmpData = null;
        if (type == Types.ASCII_STRING.getId()) {
            try {
                tmpData = new String(data, CODEPAGE);
            } catch (UnsupportedEncodingException e) {
                throw new RuntimeException("JVM Broken - core encoding " + CODEPAGE + " missing");
            }
        } else if (type == Types.UNICODE_STRING.getId()) {
            tmpData = StringUtil.getFromUnicodeLE(data);
        } else {
            throw new IllegalArgumentException("Not a string type " + type);
        }

        // Strip off the null terminator if present
        if (tmpData.endsWith("\0")) {
            tmpData = tmpData.substring(0, tmpData.length() - 1);
        }
        this.data = tmpData;
    }

    /**
     * Returns the string of a Attribute, converting as appropriate
     */
    public static String getAsString(MAPIAttribute attr) {
        if (attr == null) {
            return null;
        }
        if (attr instanceof MAPIStringAttribute) {
            return ((MAPIStringAttribute) attr).getDataString();
        }
        if (attr instanceof MAPIRtfAttribute) {
            return ((MAPIRtfAttribute) attr).getDataString();
        }

        logger.log(POILogger.WARN, "Warning, non string property found: " + attr.toString());
        return null;
    }

    public String getDataString() {
        return data;
    }

    public String toString() {
        return getProperty().toString() + " " + data;
    }
}
