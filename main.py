"""
Power BI refresh trace parser.

Reads a SQL Server Profiler trace of a Power BI semantic model refresh
(saved as a Trace XML file) and writes refresh_analysis.xlsx,
which the Power BI report uses as its data source:

    Summary Info       database, user, trace start and end,
                       busy / idle time of the Timeline
    Timeline           every processed object (partition, column, hierarchy,
                       relationship) with its thread, duration and CPU time
    System Events      Preparation, Finalization, Commit/Rollback and Lock Wait;
                       Timeline span + these durations = refresh duration
    Errors             Error events and every failed or cancelled End event
    Execution Metrics  duration, peak memory and CPU of the refresh request
    Unfinished         objects with a Progress Report Begin but no End

Usage: put exactly one .xml trace next to this script and run it.
Traces kept in subfolders (e.g. XML/) are ignored.
"""

import xml.etree.ElementTree as ET
import os
import glob
import re
import json
from datetime import datetime, timedelta
import pandas as pd

# The XML is read from, and the Excel written to, the folder of this script,
# regardless of the working directory it is run from
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "refresh_analysis.xlsx")

# ============================================================
# PROCESS NAME PREFIXES TO STRIP FROM RAW EVENT TEXT
# ============================================================
PROCESS_PREFIXES = [
    "Process AttributeHierarchy",
    "Process ColumnData",
    "Process Partition",
    "Process Relationship",
    "Process Table",
    "Process MeasureGroup",
    "Process Database",
    "PrepareTablePartitionsProcessing TableData",
    "PrepareTablePartitionsProcessing",
    "Process Index",
    "Process Data",
    "Process Clear",
    "Process Full",
    "Process Default",
    "Process Update",
    "Process Add",
    "Process"
]
# Longest first, so the most specific prefix wins
PROCESS_PREFIXES_BY_LENGTH = sorted(PROCESS_PREFIXES, key=len, reverse=True)

# End events whose Success column marks a failed operation
FAILURE_EVENTS = ('Command End', 'Progress Report End', 'Query End', 'Discover End')

# Progress Report End subclasses that are refresh phases, not object processing.
# Only Commit and Rollback are listed in System Events; the others nest inside
# Commit (Backup, commit phases) or overlap the Timeline (Sequence point),
# so they are kept out of both sheets
REFRESH_PHASES = {'6', '7', '37', '52', '57', '58'}
TRANSACTION_END = {'6': "Commit", '7': "Rollback"}

# Command End subclass of the whole refresh command; its Duration matches
# durationMs in Execution Metrics
REFRESH_COMMAND_SUBCLASS = '38'

# Markers of the refresh command text (TMSL JSON or XMLA). Discover batches
# (<Batch><Discover>) do not contain them, so they are not taken for the refresh
REFRESH_COMMAND_MARKERS = ('"refresh":', '<refresh', '<process')

TIMELINE_COLUMNS = [
    'ObjectName', 'TableName', 'Duration_Sec',
    'CPU_Duration_Sec', 'Thread_ID', 'Details', 'StartTime', 'EndTime', 'RequestID'
]
SYSTEM_EVENTS_COLUMNS = ['Action', 'Duration_Sec', 'StartTime', 'RequestID', 'TextData']
ERRORS_COLUMNS = ['ID', 'Action', 'TextData', 'Error', 'ErrorCode', 'StartTime', 'RequestID']
UNFINISHED_COLUMNS = [
    'ObjectName', 'EventSubclass', 'Details', 'StartTime', 'Open_Until_Trace_End_Sec', 'RequestID'
]
EXEC_METRICS_COLUMNS = [
    'Command Type', 'Refresh Parallelism', 'Duration Sec', 'Peak RAM GB',
    'Total CPU Sec', 'M Engine CPU Sec', 'VertiPaq CPU Sec', 'StartTime', 'EndTime', 'RequestID'
]


# ============================================================
# TEXT AND TIME HELPERS
# ============================================================
def split_name_and_process(raw_text):
    """Split a raw event line into (object_name, process_type)."""
    if not raw_text:
        return "System Process", "Unknown Process"

    clean_obj = raw_text.strip()
    process_type = ""

    for prefix in PROCESS_PREFIXES_BY_LENGTH:
        if clean_obj.startswith(prefix):
            process_type = prefix
            clean_obj = clean_obj[len(prefix):].strip()
            break

    if not clean_obj:
        clean_obj = "System Process"

    return clean_obj, process_type


def get_table_name(obj_name, process_type=""):
    """Extract the clean table name from an object name."""
    if process_type == "Process Relationship":
        return "Relationship"

    if not obj_name or obj_name == "System Process":
        return ""

    if 'H$' in obj_name:
        parts = obj_name.split('$')
        if len(parts) >= 2:
            return re.sub(r'\s*\(\d+\)', '', parts[1]).strip()

    if obj_name.startswith("'"):
        parts = obj_name.split("'")
        if len(parts) >= 3:
            return parts[1]

    if "[" in obj_name:
        return obj_name.split("[")[0].strip()

    return obj_name.strip("' ")


def resolve_object_details(clean_obj_name, process_type, fallback_text):
    """Derive (table_name, display_object_name, details) for a processed object."""
    t_name = get_table_name(clean_obj_name, process_type)

    final_obj_name = clean_obj_name[:120]
    if process_type == "Process Relationship":
        final_obj_name = f"Relationship[{final_obj_name}]"
    elif process_type.startswith("PrepareTablePartitionsProcessing"):
        final_obj_name += " (prepare)"

    raw_details = process_type if process_type else fallback_text[:100]
    if raw_details.startswith("PrepareTablePartitionsProcessing"):
        raw_details = "Process Prepare Partitions"

    return t_name, final_obj_name, raw_details


def strip_namespace(tag):
    return tag.split('}', 1)[1] if '}' in tag else tag


def clean_text(text):
    """Remove the <oii>/<pii> markers Power BI wraps around identifiers."""
    if not text:
        return ""
    return re.sub(r'</?(oii|pii)>', '', text, flags=re.IGNORECASE)


def parse_pbi_time(time_str):
    """Parse a Profiler timestamp (with up to microsecond precision) into a datetime."""
    if not time_str:
        return None

    clean_str = time_str.strip().replace('Z', '+00:00')

    # Python 3.10 accepts only 3 or 6 fractional digits and a full +HH:MM offset:
    # cut or pad the fraction to 6 digits and keep the offset as it is
    match = re.match(r'^([^.]+)\.(\d+)(.*)$', clean_str)
    if match:
        base, fraction, zone = match.groups()
        clean_str = f"{base}.{fraction[:6].ljust(6, '0')}{zone}"

    try:
        return datetime.fromisoformat(clean_str)
    except ValueError:
        return None


def parse_time_span_to_seconds(time_str):
    """Convert a Job Graph 'HH:MM:SS.fff' CPU time into seconds."""
    if not time_str:
        return 0.0
    try:
        parts = time_str.split(':')
        if len(parts) == 3:
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
    except ValueError:
        pass
    return 0.0


def ms_to_seconds(value):
    """Convert a millisecond value from a trace column into seconds (0.0 if empty)."""
    return int(value) / 1000.0 if value and value.lstrip('-').isdigit() else 0.0


def format_ts(dt, fallback=""):
    """Format a datetime as 'YYYY-MM-DD HH:MM:SS.mmm', or return fallback."""
    return dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] if dt else fallback


def find_xml_files():
    """XML files next to this script. Subfolders (e.g. XML/) are ignored."""
    found = {}
    for pattern in ("*.xml", "*.XML"):
        for path in glob.glob(os.path.join(SCRIPT_DIR, pattern)):
            # Windows glob is case-insensitive, so both patterns can return the same file
            found[os.path.normcase(path)] = path
    return sorted(found.values())


# ============================================================
# TIMELINE AND SYSTEM EVENTS BUILDERS
# ============================================================
def build_timeline_from_job_graph(job_graph_chunks):
    """Turn the Job Graph chunks (per RequestID) into Timeline rows, one per node."""
    events = []

    for r_id, chunks in job_graph_chunks.items():
        full_graph_text = "".join(chunks[i] for i in sorted(chunks.keys()) if chunks[i])

        for match in re.finditer(r'<Node\s+([^>]+)/>', full_graph_text):
            attrs_str = match.group(1)

            label_match = re.search(r'Label="([^"]+)"', attrs_str)
            slot_match = re.search(r'Slot="([^"]+)"', attrs_str)
            cpu_match = re.search(r'Cpu="([^"]+)"', attrs_str)
            start_match = re.search(r'StartedAt="([^"]+)"', attrs_str)
            end_match = re.search(r'FinishedAt="([^"]+)"', attrs_str)

            label_raw = label_match.group(1).replace('&amp;#xA;', ' ').replace('&#xA;', ' ') if label_match else ""
            slot = int(slot_match.group(1)) if slot_match else 0
            cpu_str = cpu_match.group(1) if cpu_match else "00:00:00.000"

            label = clean_text(label_raw)
            start_time = parse_pbi_time(start_match.group(1) if start_match else "")
            end_time = parse_pbi_time(end_match.group(1) if end_match else "")
            cpu_sec = parse_time_span_to_seconds(cpu_str)

            clean_obj_name, process_type = split_name_and_process(label)
            t_name, final_obj_name, raw_details = resolve_object_details(clean_obj_name, process_type, label)

            if start_time and end_time:
                dur_sec = (end_time - start_time).total_seconds()
                if dur_sec >= 0:
                    events.append({
                        'ObjectName': final_obj_name,
                        'TableName': t_name,
                        'Duration_Sec': round(dur_sec, 6),
                        'CPU_Duration_Sec': round(cpu_sec, 6),
                        'Thread_ID': f"Thread {slot + 1}",
                        'Details': raw_details,
                        'StartTime': start_time,
                        'EndTime': end_time,
                        'RequestID': r_id
                    })

    return events


def assign_fallback_threads(fallback_events):
    """No Job Graph in the trace: assign events to threads by overlap
    (greedy interval scheduling) and return them sorted by start time."""
    fallback_events.sort(key=lambda x: x['StartTime'])
    threads = []

    for ev in fallback_events:
        assigned_thread = -1
        for i, thread_free_time in enumerate(threads):
            if thread_free_time <= ev['StartTime']:
                assigned_thread = i + 1
                threads[i] = ev['EndTime']
                break
        if assigned_thread == -1:
            threads.append(ev['EndTime'])
            assigned_thread = len(threads)
        ev['Thread_ID'] = f"Thread {assigned_thread}"

    return fallback_events


def select_refresh_ops(events, refresh_command):
    """Timeline rows of the refresh request. Without a Job Graph the Timeline
    also holds other requests; only the refresh request's own operations count."""
    request_id = refresh_command['RequestID'] if refresh_command else events[0]['RequestID']
    return request_id, ([ev for ev in events if ev['RequestID'] == request_id] or events)


def compute_activity_stats(refresh_ops):
    """Wall-clock span of the Timeline split into busy time (at least one
    operation running) and idle time (gaps with nothing running).
    Sum_Object_Duration_Sec adds every operation up, so parallel work counts twice."""
    spans = sorted((ev['StartTime'].replace(tzinfo=None), ev['EndTime'].replace(tzinfo=None))
                   for ev in refresh_ops)
    busy_sec = 0.0
    cur_start, cur_end = spans[0]
    for start, end in spans[1:]:
        if start > cur_end:
            busy_sec += (cur_end - cur_start).total_seconds()
            cur_start, cur_end = start, end
        elif end > cur_end:
            cur_end = end
    busy_sec += (cur_end - cur_start).total_seconds()

    span_sec = (max(end for _, end in spans) - spans[0][0]).total_seconds()
    return {
        'Timeline_Span_Sec': round(span_sec, 6),
        'Busy_Time_Sec': round(busy_sec, 6),
        'Idle_Time_Sec': round(max(span_sec - busy_sec, 0.0), 6),
        'Sum_Object_Duration_Sec': round(sum(ev['Duration_Sec'] for ev in refresh_ops), 6),
    }


def add_refresh_phases(events, system_events_list, refresh_command, request_start_times):
    """Add Preparation and Finalization rows to System Events so that
    Preparation + Processing (Timeline span) + Commit/Rollback + Finalization
    equals the refresh command duration (and durationMs in Execution Metrics)."""
    if not events:
        return

    request_id, refresh_ops = select_refresh_ops(events, refresh_command)

    # Timestamps are compared as wall-clock values: Execution Metrics marks them
    # with 'Z' while the other events carry a local offset for the same clock
    first_op = min(ev['StartTime'] for ev in refresh_ops).replace(tzinfo=None)
    last_op = max(ev['EndTime'] for ev in refresh_ops).replace(tzinfo=None)
    # Processing (the Timeline span) is not listed: the report takes it from the Timeline
    processing_sec = (last_op - first_op).total_seconds()

    refresh_start = request_start_times.get(request_id) if refresh_command else None
    if not refresh_start:
        return

    refresh_start = refresh_start.replace(tzinfo=None)
    transaction_sec = sum(ev['Duration_Sec'] for ev in system_events_list
                          if ev['Action'] in TRANSACTION_END.values())
    preparation_sec = max((first_op - refresh_start).total_seconds(), 0.0)
    # Finalization is the remainder: the wait before Commit plus the tail after it.
    # Commit start/end are logged to the second, so the two cannot be split further
    remainder_sec = refresh_command['Duration_Sec'] - preparation_sec - processing_sec - transaction_sec
    finalization_sec = max(remainder_sec, 0.0)
    finalization_note = "Wait before Commit/Rollback and the tail after it (remainder of the refresh duration)"
    if remainder_sec < 0:
        # The measured parts add up to more than the refresh command: the Timeline or
        # Commit reach past the command end. A negative duration would be wrong, so it
        # is shown as 0 and the overshoot is kept in the text
        finalization_note = (f"Not a real duration: Preparation + Timeline + Commit/Rollback exceed the refresh "
                             f"command by {-remainder_sec:.3f} s, so Finalization is shown as 0")
        print(f"WARNING: the refresh phases exceed the refresh command duration by {-remainder_sec:.3f} s "
              f"(Finalization set to 0).")

    system_events_list.append({
        'Action': "Preparation",
        'Duration_Sec': round(preparation_sec, 6),
        'StartTime': refresh_start,
        'RequestID': request_id,
        'TextData': "From the refresh command start to the first Timeline operation"
    })
    system_events_list.append({
        'Action': "Finalization",
        'Duration_Sec': round(finalization_sec, 6),
        'StartTime': last_op,
        'RequestID': request_id,
        'TextData': finalization_note
    })


# ============================================================
# EXCEL OUTPUT
# ============================================================
def write_sheet(writer, records, sheet_name, columns, sort_by=None, ascending=False):
    """Write a list of dict records to an Excel sheet, keeping column order.
    Every column is always written (empty if no record has it), so the
    Power BI queries never miss a column. Header only when records is empty."""
    if records:
        df = pd.DataFrame(records).reindex(columns=columns)
        if sort_by:
            df = df.sort_values(by=sort_by, ascending=ascending)
    else:
        df = pd.DataFrame(columns=columns)
    df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_excel(output_file_path, summary_data, events, system_events_list, errors_list, exec_metrics_list,
                unfinished_list):
    """Write all sheets of the output workbook."""
    timeline_df = pd.DataFrame(events)
    timeline_df['Duration_Sec'] = timeline_df['Duration_Sec'].astype(float)
    timeline_df['CPU_Duration_Sec'] = timeline_df['CPU_Duration_Sec'].astype(float)
    timeline_df = timeline_df[[c for c in TIMELINE_COLUMNS if c in timeline_df.columns]]
    timeline_df = timeline_df.sort_values(by='Duration_Sec', ascending=False)

    with pd.ExcelWriter(output_file_path) as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary Info', index=False)
        timeline_df.to_excel(writer, sheet_name='Timeline', index=False)
        write_sheet(writer, system_events_list, 'System Events', SYSTEM_EVENTS_COLUMNS,
                    sort_by='StartTime', ascending=True)
        write_sheet(writer, errors_list, 'Errors', ERRORS_COLUMNS)
        write_sheet(writer, exec_metrics_list, 'Execution Metrics', EXEC_METRICS_COLUMNS,
                    sort_by='Duration Sec')
        write_sheet(writer, unfinished_list, 'Unfinished', UNFINISHED_COLUMNS,
                    sort_by='StartTime', ascending=True)


# ============================================================
# MAIN PARSER
# ============================================================
def build_refresh_analysis(xml_path, output_file_path):
    """Parse the trace at xml_path and write the Excel workbook to output_file_path."""
    fallback_events = []
    system_events_list = []

    # Whole refresh command (Command End) and request start times (Execution Metrics),
    # used to split the refresh into non-overlapping System Events rows
    refresh_command = None
    request_start_times = {}

    errors_list = []
    exec_metrics_list = []
    job_graph_chunks = {}
    # Session of each Job Graph request; events seen before the refresh command
    # is found are checked against the refresh session afterwards
    job_graph_activity = {}

    global_min_start = None
    global_max_end = None

    # Tracks the lifecycle of the main refresh command, so unrelated
    # background activity can be filtered out of the report
    main_request_id = None
    main_activity_id = None
    main_subclass = None
    main_command_ended = False

    # Progress Report Begin events still waiting for their End, keyed by request,
    # subclass, object path and start time (the same four values on both events)
    open_objects = {}

    # Captured once, for the Summary Info sheet
    report_db_name = ""
    report_user_name = ""

    for _, elem in ET.iterparse(xml_path, events=('end',)):
        if strip_namespace(elem.tag) != 'Event':
            continue

        event_name = elem.get('name')

        columns_data = {}
        for col in elem:
            if strip_namespace(col.tag) == 'Column':
                col_name = col.get('name')
                if col_name:
                    columns_data[col_name] = col.text if col.text is not None else ""

        start_str = columns_data.get('StartTime') or columns_data.get('CurrentTime', '')
        request_id = columns_data.get('RequestID', '')
        activity_id = columns_data.get('ActivityID', '')
        text_data = columns_data.get('TextData', '')
        success_status = columns_data.get('Success', '1')
        event_subclass = columns_data.get('EventSubclass')

        # Report header: the first DB / user name seen, replaced by the ones of the refresh command
        db_name = columns_data.get('DatabaseName', '')
        user_name = columns_data.get('NTUserName', '')
        if db_name and not report_db_name:
            report_db_name = db_name
        if user_name and not report_user_name:
            report_user_name = user_name

        # --------------------------------------------------------
        # DETECT THE MAIN REFRESH COMMAND AND FILTER OUT NOISE
        # --------------------------------------------------------
        # 1. Catch the start of the main refresh
        if event_name == 'Command Begin' and not main_request_id:
            text_lower = text_data.lower()
            if any(marker in text_lower for marker in REFRESH_COMMAND_MARKERS):
                main_request_id = request_id
                main_activity_id = activity_id
                main_subclass = event_subclass
                # The refresh command names the model and the account that started the refresh
                report_db_name = db_name or report_db_name
                report_user_name = user_name or report_user_name
                print(f"Main refresh command found. RequestID: {main_request_id} | ActivityID: {main_activity_id}")

        # 2. Once captured, drop unrelated background activity
        if main_activity_id and activity_id and activity_id != main_activity_id:
            elem.clear()
            if main_command_ended:
                print("Final metrics collected. New session detected, stopping file read.")
                break
            continue

        # 3. Mark the main command as finished (success or failure). The refresh
        #    request also runs inner commands (BeginTransaction, Batch, CommitTransaction)
        #    under the same RequestID, so the subclass must match too
        if (event_name == 'Command End' and not main_command_ended
                and request_id == main_request_id and event_subclass == main_subclass):
            main_command_ended = True
            print("Main command finished. Collecting remaining metrics...")

        # --------------------------------------------------------
        # TIME AND DURATION
        # --------------------------------------------------------
        start_time = parse_pbi_time(start_str)
        dur_sec = ms_to_seconds(columns_data.get('Duration', '0'))
        cpu_sec = ms_to_seconds(columns_data.get('CPUTime', '0'))

        end_time = None
        if start_time and dur_sec >= 0:
            end_time = start_time + timedelta(seconds=dur_sec)

            if global_min_start is None or start_time < global_min_start:
                global_min_start = start_time
            if global_max_end is None or end_time > global_max_end:
                global_max_end = end_time

        # --------------------------------------------------------
        # OBJECTS THAT STARTED BUT NEVER FINISHED
        # --------------------------------------------------------
        # Only the End events reach the Timeline, so an object that began and
        # then failed, was cancelled or was still running at the end of the
        # trace is invisible there. Its Begin without a matching End is kept.
        if (event_name in ('Progress Report Begin', 'Progress Report End')
                and event_subclass not in REFRESH_PHASES and columns_data.get('ObjectPath')):
            pair_key = (request_id, event_subclass, columns_data['ObjectPath'], start_str)
            if event_name == 'Progress Report End':
                open_objects.pop(pair_key, None)
            else:
                first_line = clean_text(text_data).split('\n')[0]
                open_objects[pair_key] = {
                    'ObjectName': columns_data.get('ObjectName', ''),
                    'EventSubclass': event_subclass,
                    'Details': first_line[:150],
                    'StartTime': parse_pbi_time(start_str),
                    'RequestID': request_id,
                    '_activity': activity_id
                }

        # --------------------------------------------------------
        # 1. ERRORS & FAILURES
        # --------------------------------------------------------
        # A failed refresh does not always emit an Error event, so every
        # End event with Success=0 and every rollback is logged as an error too.
        # These events still go through the sections below (Timeline, System Events).
        is_failure = event_name in FAILURE_EVENTS and success_status == '0'
        is_rollback = event_name == 'Progress Report End' and event_subclass == '7'

        if event_name in ['Error', 'Progress Report Error'] or is_failure or is_rollback:
            error_msg = columns_data.get('Error', '')
            if not error_msg and is_rollback:
                error_msg = "Transaction rolled back due to an error"
            elif not error_msg and is_failure:
                error_msg = "Command failed or was cancelled (Success=0)"

            errors_list.append({
                'ID': len(errors_list) + 1,
                'Action': f"{event_name} (Failed/Cancelled)" if (is_failure or is_rollback) else event_name,
                'TextData': text_data,
                'Error': error_msg,
                'ErrorCode': columns_data.get('ErrorCode', ''),
                'StartTime': format_ts(start_time, start_str),
                'RequestID': request_id,
                '_activity': activity_id
            })

        # --------------------------------------------------------
        # 2. EXECUTION METRICS
        # --------------------------------------------------------
        if event_name == 'Execution Metrics':
            if text_data:
                try:
                    metrics_json = json.loads(text_data)

                    duration_ms_json = float(metrics_json.get('durationMs') or 0)
                    peak_mem_kb = float(metrics_json.get('approximatePeakMemConsumptionKB') or 0)

                    m_engine_ms = float(metrics_json.get('mEngineCpuTimeMs') or metrics_json.get('queryProcessingCpuTimeMs') or 0)
                    total_cpu_ms = float(metrics_json.get('totalCpuTimeMs') or 0)
                    vertipaq_ms = float(metrics_json.get('vertipaqJobCpuTimeMs') or max(0, total_cpu_ms - m_engine_ms))

                    obj_path = metrics_json.get('objectPath', {})
                    json_obj_name = obj_path.get('table', '') if isinstance(obj_path, dict) else str(obj_path)
                    command_type = metrics_json.get('commandType', 'System Command')

                    if json_obj_name and json_obj_name != "{}":
                        display_name = f"{command_type} [{json_obj_name}]"
                    else:
                        display_name = command_type

                    time_start_json = metrics_json.get('timeStart')
                    if time_start_json:
                        request_start_times[request_id] = parse_pbi_time(time_start_json)

                    # Short requests with no memory or CPU usage are noise
                    if not (duration_ms_json < 1000 and peak_mem_kb == 0 and total_cpu_ms == 0):
                        time_end_json = metrics_json.get('timeEnd')
                        exec_metrics_list.append({
                            'Command Type': display_name,
                            'Refresh Parallelism': metrics_json.get('refreshParallelism', ''),
                            'Duration Sec': round(duration_ms_json / 1000.0, 4),
                            'Peak RAM GB': round(peak_mem_kb / (1024 * 1024), 4),
                            'Total CPU Sec': round(total_cpu_ms / 1000.0, 4),
                            'M Engine CPU Sec': round(m_engine_ms / 1000.0, 4),
                            'VertiPaq CPU Sec': round(vertipaq_ms / 1000.0, 4),
                            'StartTime': time_start_json.replace('T', ' ').replace('Z', '') if time_start_json else format_ts(start_time, start_str),
                            'EndTime': time_end_json.replace('T', ' ').replace('Z', '') if time_end_json else format_ts(end_time),
                            'RequestID': request_id
                        })
                except (ValueError, TypeError, AttributeError):
                    # Malformed or unexpected metrics JSON: skip this event
                    pass

        # --------------------------------------------------------
        # 3. LOCK WAITS (go to System Events)
        # --------------------------------------------------------
        # A lock wait is logged as a single point in time, so its wait time
        # is not part of the refresh duration. Deadlocks are not tracked:
        # the cancelled command already reaches Errors via Success=0
        elif event_name == 'Lock Wait':
            # Lock Wait has no Duration or EndTime columns. The WaitTime in its
            # TextData is left there: its unit is undocumented and a wait that long
            # does not fit the traced timeline, so it is not reported as a duration
            system_events_list.append({
                'Action': event_name,
                'Duration_Sec': None,
                'StartTime': start_time or start_str,
                'RequestID': request_id,
                'TextData': clean_text(text_data),
                '_activity': activity_id
            })

        # --------------------------------------------------------
        # 4. JOB GRAPH (accumulate chunks per RequestID)
        # --------------------------------------------------------
        elif event_name == 'Job Graph' and event_subclass in ('1', '2'):
            int_data_raw = columns_data.get('IntegerData', '0')
            chunk_id = int(int_data_raw) if int_data_raw and int_data_raw.lstrip('-').isdigit() else 0

            job_graph_chunks.setdefault(request_id, {})[chunk_id] = text_data
            job_graph_activity.setdefault(request_id, activity_id)

        # --------------------------------------------------------
        # 5. EVENTS TIMELINE
        # --------------------------------------------------------
        elif event_name in ['Progress Report End', 'Query End', 'Command End', 'Discover End']:
            clean_text_data = clean_text(text_data)

            if start_time and dur_sec >= 0:
                first_line = clean_text_data.split('\n')[0] if clean_text_data else ""

                # Refresh phases and the refresh command itself stay off the Timeline
                if event_name == 'Command End' and event_subclass == REFRESH_COMMAND_SUBCLASS:
                    refresh_command = {'RequestID': request_id, 'Duration_Sec': dur_sec}
                elif event_name == 'Progress Report End' and event_subclass in REFRESH_PHASES:
                    if event_subclass in TRANSACTION_END:
                        system_events_list.append({
                            'Action': TRANSACTION_END[event_subclass],
                            'Duration_Sec': round(dur_sec, 6),
                            'StartTime': start_time,
                            'RequestID': request_id,
                            'TextData': clean_text_data,
                            '_activity': activity_id
                        })
                else:
                    clean_obj_name, process_type = split_name_and_process(first_line)
                    if first_line.strip().startswith('<'):
                        final_obj_name = "XML Command"
                        t_name = ""
                        raw_details = clean_text_data.replace('\n', ' ')[:100]
                    else:
                        t_name, final_obj_name, raw_details = resolve_object_details(
                            clean_obj_name, process_type, clean_text_data.replace('\n', ' ')
                        )

                    fallback_events.append({
                        'ObjectName': final_obj_name,
                        'TableName': t_name,
                        'Duration_Sec': round(dur_sec, 6),
                        'CPU_Duration_Sec': round(cpu_sec, 6),
                        'Thread_ID': '',
                        'Details': raw_details,
                        'StartTime': start_time,
                        'EndTime': end_time,
                        'RequestID': request_id,
                        '_activity': activity_id
                    })

        elem.clear()

    # Events read before the refresh command was found were not filtered by
    # session yet: keep only those of the refresh session (or without ActivityID)
    if main_activity_id:
        def in_refresh_session(aid):
            return not aid or aid == main_activity_id

        errors_list = [e for e in errors_list if in_refresh_session(e['_activity'])]
        system_events_list = [e for e in system_events_list if in_refresh_session(e['_activity'])]
        fallback_events = [e for e in fallback_events if in_refresh_session(e['_activity'])]
        job_graph_chunks = {r: c for r, c in job_graph_chunks.items()
                            if in_refresh_session(job_graph_activity.get(r, ''))}
        open_objects = {k: v for k, v in open_objects.items() if in_refresh_session(v['_activity'])}
        for number, error in enumerate(errors_list, start=1):
            error['ID'] = number

    # --------------------------------------------------------
    # TIMELINE: from the Job Graph when present, otherwise from End events
    # --------------------------------------------------------
    if job_graph_chunks:
        events = build_timeline_from_job_graph(job_graph_chunks)
    else:
        events = assign_fallback_threads(fallback_events)

    if not events:
        print("No events found. Excel file was not updated.")
        return

    events.sort(key=lambda x: x['StartTime'])

    add_refresh_phases(events, system_events_list, refresh_command, request_start_times)

    # Execution Metrics keeps only the refresh request; the small service requests
    # around it (ActAs, Discover batches, system commands) are dropped
    refresh_ids = {refresh_command['RequestID']} if refresh_command else {ev['RequestID'] for ev in events}
    exec_metrics_list = [em for em in exec_metrics_list if em['RequestID'] in refresh_ids]

    if global_min_start and global_max_end:
        total_refresh_sec = round((global_max_end - global_min_start).total_seconds(), 6)
    else:
        total_refresh_sec = 0.0

    _, refresh_ops = select_refresh_ops(events, refresh_command)
    activity_stats = compute_activity_stats(refresh_ops)

    unfinished_list = []
    if global_max_end:
        for obj in sorted(open_objects.values(), key=lambda o: o['StartTime'] or datetime.min):
            if obj['StartTime']:
                obj['Open_Until_Trace_End_Sec'] = round(
                    max((global_max_end - obj['StartTime']).total_seconds(), 0.0), 3)
            unfinished_list.append(obj)

    print("\n======================================")
    print(f"DATABASE:      {report_db_name}")
    print(f"USER:          {report_user_name}")
    print(f"TOTAL REFRESH: {total_refresh_sec} sec")
    print("--------------------------------------")
    print(f"Errors / Failures: {len(errors_list)}")
    print(f"Execution Metrics: {len(exec_metrics_list)}")
    print(f"System Events:     {len(system_events_list)}")
    print(f"Unfinished objects: {len(unfinished_list)}")
    print(f"Timeline: span {activity_stats['Timeline_Span_Sec']} s = busy {activity_stats['Busy_Time_Sec']} s "
          f"+ idle {activity_stats['Idle_Time_Sec']} s")
    print("======================================\n")

    # Convert datetimes to display strings for both event lists
    for record_list in (events, system_events_list, unfinished_list):
        for ev in record_list:
            for key in ('StartTime', 'EndTime'):
                if isinstance(ev.get(key), datetime):
                    ev[key] = format_ts(ev[key])

    summary_data = [{
        'Database': report_db_name,
        'User': report_user_name,
        'StartTime': format_ts(global_min_start),
        'EndTime': format_ts(global_max_end),
        **activity_stats
    }]

    write_excel(output_file_path, summary_data, events, system_events_list, errors_list, exec_metrics_list,
                unfinished_list)
    print(f"File successfully created: {output_file_path}")


def main():
    xml_files = find_xml_files()
    if not xml_files:
        print("No XML files found.")
    elif len(xml_files) > 1:
        print("More than one XML found next to main.py. Leave only the one to parse:")
        for path in xml_files:
            print(f"  {os.path.basename(path)}")
    else:
        print(f"Using XML: {xml_files[0]}")
        build_refresh_analysis(xml_files[0], OUTPUT_FILE)


if __name__ == "__main__":
    main()
