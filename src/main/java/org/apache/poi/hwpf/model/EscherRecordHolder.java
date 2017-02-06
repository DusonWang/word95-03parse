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

package org.apache.poi.hwpf.model;

import org.apache.poi.ddf.DefaultEscherRecordFactory;
import org.apache.poi.ddf.EscherContainerRecord;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.ddf.EscherRecordFactory;
import org.apache.poi.util.Internal;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Based on AbstractEscherRecordHolder from HSSF.
 *
 * @author Squeeself
 */
@Internal
public final class EscherRecordHolder {
    private final ArrayList<EscherRecord> escherRecords;

    public EscherRecordHolder() {
        escherRecords = new ArrayList<>();
    }

    public EscherRecordHolder(byte[] data, int offset, int size) {
        this();
        fillEscherRecords(data, offset, size);
    }

    private static EscherRecord findFirstWithId(short id, List<EscherRecord> records) {
        // Check at our level
        for (EscherRecord r : records) {
            if (r.getRecordId() == id) {
                return r;
            }
        }

        // Then check our children in turn
        for (EscherRecord r : records) {
            if (r.isContainerRecord()) {
                EscherRecord found = findFirstWithId(id, r.getChildRecords());
                if (found != null) {
                    return found;
                }
            }
        }

        // Not found in this lot
        return null;
    }

    private void fillEscherRecords(byte[] data, int offset, int size) {
        EscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
        int pos = offset;
        while (pos < offset + size) {
            EscherRecord r = recordFactory.createRecord(data, pos);
            escherRecords.add(r);
            int bytesRead = r.fillFields(data, pos, recordFactory);
            pos += bytesRead + 1; // There is an empty byte between each top-level record in a Word doc
        }
    }

    public List<EscherRecord> getEscherRecords() {
        return escherRecords;
    }

    public String toString() {
        StringBuilder buffer = new StringBuilder();

        if (escherRecords.size() == 0) {
            buffer.append("No Escher Records Decoded").append("\n");
        }
        for (EscherRecord r : escherRecords) {
            buffer.append(r.toString());
        }
        return buffer.toString();
    }

    /**
     * If we have a EscherContainerRecord as one of our
     * children (and most top level escher holders do),
     * then return that.
     */
    public EscherContainerRecord getEscherContainer() {
        for (EscherRecord er : escherRecords) {
            if (er instanceof EscherContainerRecord) {
                return (EscherContainerRecord) er;
            }
        }
        return null;
    }

    /**
     * Descends into all our children, returning the
     * first EscherRecord with the given id, or null
     * if none found
     */
    public EscherRecord findFirstWithId(short id) {
        return findFirstWithId(id, getEscherRecords());
    }

    public List<? extends EscherContainerRecord> getDgContainers() {
        List<EscherContainerRecord> dgContainers = new ArrayList<>(
                1);
        dgContainers.addAll(getEscherRecords().stream().filter(escherRecord -> escherRecord.getRecordId() == (short) 0xF002).map(escherRecord -> (EscherContainerRecord) escherRecord).collect(Collectors.toList()));
        return dgContainers;
    }

    public List<? extends EscherContainerRecord> getDggContainers() {
        List<EscherContainerRecord> dggContainers = new ArrayList<>(
                1);
        dggContainers.addAll(getEscherRecords().stream().filter(escherRecord -> escherRecord.getRecordId() == (short) 0xF000).map(escherRecord -> (EscherContainerRecord) escherRecord).collect(Collectors.toList()));
        return dggContainers;
    }

    public List<? extends EscherContainerRecord> getBStoreContainers() {
        List<EscherContainerRecord> bStoreContainers = new ArrayList<>(
                1);
        for (EscherContainerRecord dggContainer : getDggContainers()) {
            bStoreContainers.addAll(dggContainer.getChildRecords().stream().filter(escherRecord -> escherRecord.getRecordId() == (short) 0xF001).map(escherRecord -> (EscherContainerRecord) escherRecord).collect(Collectors.toList()));
        }
        return bStoreContainers;
    }

    public List<? extends EscherContainerRecord> getSpgrContainers() {
        List<EscherContainerRecord> spgrContainers = new ArrayList<>(
                1);
        for (EscherContainerRecord dgContainer : getDgContainers()) {
            spgrContainers.addAll(dgContainer.getChildRecords().stream().filter(escherRecord -> escherRecord.getRecordId() == (short) 0xF003).map(escherRecord -> (EscherContainerRecord) escherRecord).collect(Collectors.toList()));
        }
        return spgrContainers;
    }

    public List<? extends EscherContainerRecord> getSpContainers() {
        List<EscherContainerRecord> spContainers = new ArrayList<>(
                1);
        for (EscherContainerRecord spgrContainer : getSpgrContainers()) {
            spContainers.addAll(spgrContainer.getChildRecords().stream().filter(escherRecord -> escherRecord.getRecordId() == (short) 0xF004).map(escherRecord -> (EscherContainerRecord) escherRecord).collect(Collectors.toList()));
        }
        return spContainers;
    }
}
