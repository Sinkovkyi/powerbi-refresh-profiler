"""Build the synthetic Profiler traces in tests/fixtures (no real data):
sample_refresh_trace.xml (one refresh with a Job Graph) and two_attempts_trace.xml (a failed attempt, then a
successful one, without a Job Graph).

The fixture is UTF-16 like a real Profiler export, so Git shows it as binary: change this script,
not the XML, then run it and the tests:  python tests/make_fixture.py && python -m unittest discover tests
"""
import json
import os
import xml.etree.ElementTree as ET

NS = "http://tempuri.org/TracePersistence.xsd"
A = "AAAAAAAA-0000-4000-8000-000000000001"   # refresh session
B = "BBBBBBBB-0000-4000-8000-000000000002"   # another session
R0 = "0000A0A0-0000-4000-8000-000000000010"  # service batch before the refresh
R1 = "11111111-0000-4000-8000-000000000011"  # refresh request
DB = "Contoso Sales"
USER = "Power BI Service (analyst@contoso.com)"

root = ET.Element("TraceData", xmlns=NS)
hdr = ET.SubElement(root, "Header")
ET.SubElement(hdr, "TraceProvider", name="Microsoft Analysis Services", MajorVersion="17", MinorVersion="0",
              BuildNumber="88")
ET.SubElement(hdr, "ServerInformation", name="powerbi://api.powerbi.com/v1.0/myorg/Demo Workspace")
events = ET.SubElement(root, "Events")
ids = {"Command Begin": "15", "Command End": "16", "Progress Report Begin": "5", "Progress Report End": "6",
       "Query End": "10", "Error": "17",
       "Lock Wait": "54", "Job Graph": "134", "Execution Metrics": "136"}


def event(name, **cols):
    e = ET.SubElement(events, "Event", id=ids[name], name=name)
    for key, value in cols.items():
        c = ET.SubElement(e, "Column", name=key)
        c.text = str(value)


# 1. A failed query of another session before the refresh starts: must not reach Errors
event("Query End", EventSubclass="3", StartTime="2026-09-08T09:59:58+02:00", Duration="100", CPUTime="0",
      Success="0", Error="-1055784777", ActivityID=B, RequestID="22222222-0000-4000-8000-000000000020",
      DatabaseName=DB, NTUserName="viewer@contoso.com", TextData="EVALUATE 'Sales'")

# 2. Service batch of the refresh session: not the refresh command
event("Command Begin", EventSubclass="12", StartTime="2026-09-08T10:00:00+02:00", ActivityID=A, RequestID=R0,
      DatabaseName=DB, NTUserName=USER, TextData='<Batch Transaction="false"><Alter /></Batch>')
event("Command End", EventSubclass="12", StartTime="2026-09-08T10:00:00+02:00", Duration="20", CPUTime="0",
      Success="1", ActivityID=A, RequestID=R0, DatabaseName=DB, NTUserName=USER,
      TextData='<Batch Transaction="false"><Alter /></Batch>')

# 3. The refresh command (TMSL)
tmsl = json.dumps({"sequence": {"maxParallelism": 2, "operations": [
    {"refresh": {"type": "full", "objects": [{"database": DB}]}}]}}, indent=2)
event("Command Begin", EventSubclass="38", StartTime="2026-09-08T10:00:00+02:00", ActivityID=A, RequestID=R1,
      DatabaseName=DB, NTUserName=USER, TextData=tmsl)

# 4. Lock wait of the refresh
event("Lock Wait", EventSubclass="0", StartTime="2026-09-08T10:00:00+02:00", ActivityID=A, RequestID=R1,
      DatabaseName=DB, NTUserName=USER,
      TextData="LockType=Exclusive, ObjectPath=Contoso Sales/Model/Tables/Sales, BlockedSPID=101, "
               "BlockingSPID=202, WaitTime=100")

# 5. Job Graph (subclass 2 = executed graph), split into chunks like Profiler does
nodes = [
    ("1", "PrepareTablePartitionsProcessing TableData&#xA;Sales", "0", "00:00:00.100",
     "2026-09-08T10:00:01.000Z", "2026-09-08T10:00:01.200Z"),
    ("2", "Process Partition&#xA;Sales[Sales]", "0", "00:00:01.000",
     "2026-09-08T10:00:01.200Z", "2026-09-08T10:00:05.000Z"),
    ("3", "Process AttributeHierarchy&#xA;Sales[CustomerKey]", "1", "00:00:01.900",
     "2026-09-08T10:00:05.000Z", "2026-09-08T10:00:07.000Z"),
    ("4", "Process Relationship&#xA;0f0e0d0c-0000-4000-8000-000000000001", "0", "00:00:00.400",
     "2026-09-08T10:00:07.000Z", "2026-09-08T10:00:07.500Z"),
]
dgml = ('<DirectedGraph Background="#FFFFFF" GraphDirection="TopToBottom" Layout="Sugiyama" '
        'Title="Job Graph G1" xmlns="http://schemas.microsoft.com/vs/2009/dgml"><Nodes>')
for nid, label, slot, cpu, started, finished in nodes:
    dgml += (f'<Node Id="{nid}" Label="{label}" Slot="{slot}" Block="00:00:00.000"  Wait="00:00:00.000"  '
             f'Run="00:00:00.000"  Cpu="{cpu}"  CreatedAt="2026-09-08T10:00:00.900Z"  '
             f'RunnableAt="{started}"  StartedAt="{started}"  FinishedAt="{finished}"  Category="Success"  />')
dgml += '</Nodes><Links><Link Source="2" Target="1" Category="Standard" /></Links></DirectedGraph>'
half = len(dgml) // 2
event("Job Graph", EventSubclass="2", IntegerData="0", CurrentTime="2026-09-08T10:00:08+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER,
      TextData=f'<Metadata Length="{len(dgml)}" AdditionalEvents="2" />')
event("Job Graph", EventSubclass="2", IntegerData="1", CurrentTime="2026-09-08T10:00:08+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, TextData=dgml[:half])
event("Job Graph", EventSubclass="2", IntegerData="2", CurrentTime="2026-09-08T10:00:08+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, TextData=dgml[half:])

# 5b. Progress Report Begin / End pairs. The partition (subclass 59) finishes; the relationship
# build (subclass 46) never gets an End, like a relationship that was still running when the trace stopped
event("Progress Report Begin", EventSubclass="59", StartTime="2026-09-08T10:00:01+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, ObjectName="Sales", ObjectPath=f"{DB}.Model.Sales.Sales",
      TextData="Started processing partition 'Sales'.")
event("Progress Report End", EventSubclass="59", StartTime="2026-09-08T10:00:01+02:00", Duration="3800",
      CPUTime="1000", Success="1", Error="0", ActivityID=A, RequestID=R1, DatabaseName=DB, NTUserName=USER,
      ObjectName="Sales", ObjectPath=f"{DB}.Model.Sales.Sales", TextData="Finished processing partition 'Sales'.")
event("Progress Report Begin", EventSubclass="46", StartTime="2026-09-08T10:00:07+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, ObjectName="Sales", ObjectPath=f"{DB}.Model.Sales.Rel1",
      TextData="Started building relationship 'Sales'[CustomerKey] -> 'Customer'[CustomerKey].")

# 5c. Two errors of the refresh session, written in the order they were logged: the later one first.
# The Errors sheet must put the earlier one on top (ID 1)
event("Error", EventSubclass="1", Error="3238395905", StartTime="2026-09-08T10:00:06+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, TextData="Second error: the operation was cancelled.")
event("Error", EventSubclass="1", Error="3238395904", StartTime="2026-09-08T10:00:02+02:00", ActivityID=A,
      RequestID=R1, DatabaseName=DB, NTUserName=USER, TextData="First error: out of memory.")

# 6. Refresh phases: Sequence point overlaps the timeline (not listed), Commit is listed
event("Progress Report End", EventSubclass="58", StartTime="2026-09-08T10:00:01+02:00", Duration="6500",
      CPUTime="0", Success="1", Error="0", ActivityID=A, RequestID=R1, DatabaseName=DB, NTUserName=USER,
      TextData="Finished sequence point algorithm.")
event("Progress Report End", EventSubclass="6", StartTime="2026-09-08T10:00:08+02:00", Duration="1500",
      CPUTime="0", Success="1", Error="0", ActivityID=A, RequestID=R1, DatabaseName=DB, NTUserName=USER)

# 7. End of the refresh command and its metrics
event("Command End", EventSubclass="38", StartTime="2026-09-08T10:00:00+02:00", Duration="9000", CPUTime="4000",
      Success="1", ActivityID=A, RequestID=R1, DatabaseName=DB, NTUserName=USER, TextData=tmsl)
metrics = {"timeStart": "2026-09-08T10:00:00.500Z", "timeEnd": "2026-09-08T10:00:09.500Z", "durationMs": 9000,
           "vertipaqJobCpuTimeMs": 3900, "mEngineCpuTimeMs": 100, "totalCpuTimeMs": 4000,
           "approximatePeakMemConsumptionKB": 1048576, "commandType": "Batch", "refreshParallelism": 2}
event("Execution Metrics", CurrentTime="2026-09-08T10:00:09+02:00", ActivityID=A, RequestID=R1,
      DatabaseName=DB, TextData=json.dumps(metrics, indent=1))
event("Execution Metrics", CurrentTime="2026-09-08T10:00:09+02:00", ActivityID=A, RequestID=R0, DatabaseName=DB,
      TextData=json.dumps({"timeStart": "2026-09-08T10:00:00.000Z", "timeEnd": "2026-09-08T10:00:00.020Z",
                           "durationMs": 20, "commandType": "Batch"}))

# 8. Another session after the refresh: stops the reading
event("Query End", EventSubclass="3", StartTime="2026-09-08T10:00:10+02:00", Duration="50", CPUTime="0",
      Success="1", ActivityID=B, RequestID="33333333-0000-4000-8000-000000000030", DatabaseName=DB,
      NTUserName="viewer@contoso.com", TextData="EVALUATE 'Sales'")



def write(tree_root, name):
    ET.indent(tree_root)
    xml = '<?xml version="1.0" encoding="utf-16"?>\n' + ET.tostring(tree_root, encoding="unicode")
    output = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", name)
    with open(output, "w", encoding="utf-16", newline="\n") as f:
        f.write(xml)
    print(f"Written: {output}")


write(root, "sample_refresh_trace.xml")

# ------------------------------------------------------------------
# two_attempts_trace.xml: the same session runs the refresh twice. Attempt 1 fails, is rolled back and logs
# no Execution Metrics; attempt 2 succeeds. No Job Graph, so the Timeline comes from Progress Report End
# ------------------------------------------------------------------
root = ET.Element("TraceData", xmlns=NS)
events = ET.SubElement(root, "Events")
R_FAIL = "AAAA1111-0000-4000-8000-000000000101"
R_OK = "BBBB2222-0000-4000-8000-000000000102"
for rid, t0, ok in ((R_FAIL, "10:00:00", False), (R_OK, "10:01:00", True)):
    event("Command Begin", EventSubclass="38", StartTime=f"2026-09-08T{t0}+02:00", ActivityID=A, RequestID=rid,
          DatabaseName=DB, NTUserName=USER, TextData=tmsl)
# attempt 1: partition 10:00:01-10:00:03, Rollback 0.2 s, command 4 s, failed
event("Progress Report End", EventSubclass="59", StartTime="2026-09-08T10:00:01+02:00", Duration="2000",
      CPUTime="500", Success="1", Error="0", ActivityID=A, RequestID=R_FAIL, DatabaseName=DB, NTUserName=USER,
      ObjectName="Sales", TextData="Finished processing the 'Sales' partition.")
event("Progress Report End", EventSubclass="7", StartTime="2026-09-08T10:00:03+02:00", Duration="200",
      CPUTime="0", Success="1", Error="0", ActivityID=A, RequestID=R_FAIL, DatabaseName=DB, NTUserName=USER)
event("Command End", EventSubclass="38", StartTime="2026-09-08T10:00:00+02:00", Duration="4000", CPUTime="500",
      Success="0", Error="-1055784777", ActivityID=A, RequestID=R_FAIL, DatabaseName=DB, NTUserName=USER,
      TextData=tmsl)
# attempt 2: partition 10:01:02-10:01:05, Commit 0.5 s, command 6 s, Execution Metrics from 10:01:00.500
event("Progress Report End", EventSubclass="59", StartTime="2026-09-08T10:01:02+02:00", Duration="3000",
      CPUTime="900", Success="1", Error="0", ActivityID=A, RequestID=R_OK, DatabaseName=DB, NTUserName=USER,
      ObjectName="Sales", TextData="Finished processing the 'Sales' partition.")
event("Progress Report End", EventSubclass="6", StartTime="2026-09-08T10:01:05+02:00", Duration="500",
      CPUTime="0", Success="1", Error="0", ActivityID=A, RequestID=R_OK, DatabaseName=DB, NTUserName=USER)
event("Command End", EventSubclass="38", StartTime="2026-09-08T10:01:00+02:00", Duration="6000", CPUTime="900",
      Success="1", ActivityID=A, RequestID=R_OK, DatabaseName=DB, NTUserName=USER, TextData=tmsl)
event("Execution Metrics", CurrentTime="2026-09-08T10:01:07+02:00", ActivityID=A, RequestID=R_OK, DatabaseName=DB,
      TextData=json.dumps({"timeStart": "2026-09-08T10:01:00.500Z", "timeEnd": "2026-09-08T10:01:06.500Z",
                           "durationMs": 6000, "totalCpuTimeMs": 900, "approximatePeakMemConsumptionKB": 524288,
                           "commandType": "Batch", "refreshParallelism": 1}))
write(root, "two_attempts_trace.xml")
