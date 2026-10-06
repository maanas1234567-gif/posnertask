import { initJsPsych } from 'https://cdn.jsdelivr.net/npm/jspsych@8.2.1/+esm';
import htmlKeyboardResponse from 'https://cdn.jsdelivr.net/npm/@jspsych/plugin-html-keyboard-response@2.1.0/+esm';
import preloadPlugin from 'https://cdn.jsdelivr.net/npm/@jspsych/plugin-preload@2.1.0/+esm';
import * as XLSX from 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';

const EXPERIMENT_VERSION = '1.1.0';
// Remote upload is an optional future integration; local Excel export is always enabled.
const REMOTE_DATA_CONFIG = {
  enabled: false,
  provider: 'DataPipe',
  experimentId: ''
};

function createSessionId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return `session-${Date.now()}-${Math.random().toString(36).slice(2, 12)}`;
}

const sessionMetadata = {
  participant_id: createSessionId(),
  session_id: createSessionId(),
  experiment_version: EXPERIMENT_VERSION,
  experiment_start_timestamp: new Date().toISOString(),
  browser_user_agent: navigator.userAgent,
  screen_width: screen.width,
  screen_height: screen.height,
  viewport_width: window.innerWidth,
  viewport_height: window.innerHeight
};

document.getElementById('jspsych-target').innerHTML = '<section class="instruction"><h1>Loading experiment...</h1></section>';

const jsPsych = initJsPsych({
  display_element: 'jspsych-target',
  on_trial_finish: (data) => {
    data.participant_response = data.response ?? '';
  },
  on_finish: () => {
    const rawCsv = jsPsych.data.get().csv();
    console.log('Experiment data (raw jsPsych CSV):', rawCsv);

    const exportColumns = [
      ['Row Number', 'trial_number'],
      ['Phase', 'phase'],
      ['Event', 'trial_type'],
      ['Task Code', 'task'],
      ['Trial in Block', 'trial_index'],
      ['Recognition Image', 'image_id'],
      ['Center Target Image', 'center_image_id'],
      ['Matching Target Image', 'match_image_id'],
      ['Distractor Images', 'distractor_image_ids'],
      ['Cue Direction', 'cue_direction'],
      ['Cue Validity', 'cue_validity'],
      ['Target Location', 'match_location'],
      ['Correct Key', 'correct_response'],
      ['Response Key', 'response_key'],
      ['Response Meaning', 'response_meaning'],
      ['Response Outcome', 'response_outcome'],
      ['Response Time (ms)', 'rt_ms'],
      ['Duration (ms)', 'duration_ms'],
      ['Memory Status', 'image_status'],
      ['Accuracy Code', 'accuracy'],
      ['Source Phase 1 Trial', 'source_phase1_trial'],
      ['Source Cue Validity', 'source_cue_validity'],
      ['Source Target Location', 'source_target_location'],
      ['Source Cue Direction', 'source_cue_direction'],
      ['Source Phase 1 Response', 'source_phase1_response'],
      ['Source Phase 1 Correct', 'source_phase1_correct'],
      ['Source Phase 1 RT (ms)', 'source_phase1_rt'],
      ['Participant ID', 'participant_id'],
      ['Session ID', 'session_id'],
      ['Experiment Version', 'experiment_version'],
      ['Experiment Start Timestamp', 'experiment_start_timestamp'],
      ['Browser User Agent', 'browser_user_agent'],
      ['Screen Width', 'screen_width'],
      ['Screen Height', 'screen_height'],
      ['Viewport Width', 'viewport_width'],
      ['Viewport Height', 'viewport_height']
    ];
    const rawTrials = jsPsych.data.get().values();
    const responseTasks = new Set(['practice_matching', 'phase1_matching', 'recognition']);
    const exportRows = jsPsych.data.get().values().map((trial, index) => ({
      trial_number: index + 1,
      phase: trial.phase ?? '',
      trial_type: trial.trial_type ?? '',
      task: trial.task ?? '',
      trial_index: trial.trial_index ?? '',
      image_id: trial.image_id ?? '',
      center_image_id: trial.center_image_id ?? '',
      match_image_id: trial.match_image_id ?? '',
      distractor_image_ids: trial.distractor_image_ids ?? '',
      cue_direction: trial.cue_direction ?? '',
      cue_validity: trial.validity ?? '',
      match_location: trial.match_location ?? '',
      correct_response: trial.correct_response ?? '',
      response_key: trial.participant_response || '',
      response_meaning: describeResponse(trial.participant_response, trial.task),
      response_outcome: !responseTasks.has(trial.task)
        ? 'Not applicable'
        : !trial.participant_response
          ? 'No response'
          : trial.accuracy === 1
            ? 'Correct'
            : 'Incorrect',
      accuracy: trial.accuracy ?? '',
      rt_ms: trial.rt ?? '',
      duration_ms: trial.duration_ms ?? '',
      image_status: trial.image_status ?? '',
      source_phase1_trial: trial.source_phase1_trial ?? '',
      source_cue_validity: trial.source_cue_validity ?? '',
      source_target_location: trial.source_target_location ?? '',
      source_cue_direction: trial.source_cue_direction ?? '',
      source_phase1_response: trial.source_phase1_response ?? '',
      source_phase1_correct: trial.source_phase1_correct ?? '',
      source_phase1_rt: trial.source_phase1_rt ?? '',
      participant_id: trial.participant_id ?? '',
      session_id: trial.session_id ?? '',
      experiment_version: trial.experiment_version ?? '',
      experiment_start_timestamp: trial.experiment_start_timestamp ?? '',
      browser_user_agent: trial.browser_user_agent ?? '',
      screen_width: trial.screen_width ?? '',
      screen_height: trial.screen_height ?? '',
      viewport_width: trial.viewport_width ?? '',
      viewport_height: trial.viewport_height ?? ''
    }));
    const resultRows = exportRows.map((row) => Object.fromEntries(
      exportColumns.map(([label, key]) => [label, row[key]])
    ));
    const responseRows = exportRows.filter((row) => responseTasks.has(row.task));
    const phase1Rows = exportRows.filter((row) => row.task === 'phase1_matching');
    const recognitionRows = exportRows.filter((row) => row.task === 'recognition');
    const answeredRows = responseRows.filter((row) => row.response_key !== '');
    const correctRows = answeredRows.filter((row) => row.response_outcome === 'Correct');
    const correctPhase1Rows = correctRows.filter((row) => row.task === 'phase1_matching');
    const correctRecognitionRows = correctRows.filter((row) => row.task === 'recognition');
    const validPhase1Rows = phase1Rows.filter((row) => row.cue_validity === 'valid');
    const invalidPhase1Rows = phase1Rows.filter((row) => row.cue_validity === 'invalid');
    const validCorrectRows = validPhase1Rows.filter((row) => row.response_outcome === 'Correct');
    const invalidCorrectRows = invalidPhase1Rows.filter((row) => row.response_outcome === 'Correct');
    const oldValidRows = recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'valid');
    const oldInvalidRows = recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'invalid');
    const oldTrials = recognitionRows.filter((row) => row.image_status === 'old');
    const newTrials = recognitionRows.filter((row) => row.image_status === 'new');
    const hitCount = oldTrials.filter((row) => row.response_meaning === 'Old').length;
    const falseAlarmCount = newTrials.filter((row) => row.response_meaning === 'Old').length;
    const correctedHitRate = (hitCount + 0.5) / (oldTrials.length + 1);
    const correctedFalseAlarmRate = (falseAlarmCount + 0.5) / (newTrials.length + 1);
    const accuracyRate = (rows) => rows.length
      ? rows.filter((row) => row.response_outcome === 'Correct').length / rows.length
      : '';
    const phase1ImageIds = phase1Rows.flatMap((row) => [
      row.center_image_id,
      ...row.distractor_image_ids.split(',').filter(Boolean)
    ]);
    const oldRecognitionIds = recognitionRows.filter((row) => row.image_status === 'old').map((row) => row.image_id);
    const newRecognitionIds = recognitionRows.filter((row) => row.image_status === 'new').map((row) => row.image_id);
    const responseRate = (rows) => {
      const answered = rows.filter((row) => row.response_key !== '');
      return answered.length ? correctRowsFor(answered).length / answered.length : '';
    };
    const correctRowsFor = (rows) => rows.filter((row) => row.response_outcome === 'Correct');
    const meanRt = (rows) => {
      const values = rows.map((row) => Number(row.rt_ms)).filter((value) => Number.isFinite(value) && value > 0);
      return values.length ? Math.round(values.reduce((sum, value) => sum + value, 0) / values.length) : '';
    };
    const validMeanRt = meanRt(validCorrectRows);
    const invalidMeanRt = meanRt(invalidCorrectRows);
    const summaryRows = [
      { Section: 'Recorded data', Measure: 'All recorded rows (includes instructions and timed events)', Value: rawTrials.length },
      { Section: 'Recorded data', Measure: 'Practice response trials', Value: exportRows.filter((row) => row.task === 'practice_matching').length },
      { Section: 'Recorded data', Measure: 'Phase 1 response trials', Value: phase1Rows.length },
      { Section: 'Recorded data', Measure: 'Phase 2 recognition trials', Value: recognitionRows.length },
      { Section: 'Responses', Measure: 'Responses expected', Value: responseRows.length },
      { Section: 'Responses', Measure: 'Responses received', Value: responseRows.filter((row) => row.response_key !== '').length },
      { Section: 'Responses', Measure: 'No response', Value: responseRows.filter((row) => row.response_key === '').length },
      { Section: 'Responses', Measure: 'Correct responses', Value: responseRows.filter((row) => row.response_outcome === 'Correct').length },
      { Section: 'Responses', Measure: 'Incorrect responses', Value: responseRows.filter((row) => row.response_outcome === 'Incorrect').length },
      { Section: 'Performance', Measure: 'Phase 1 accuracy (% of answered matching trials)', Value: responseRate(phase1Rows) },
      { Section: 'Performance', Measure: 'Phase 1 mean RT on correct trials (ms)', Value: meanRt(correctPhase1Rows) },
      { Section: 'Performance', Measure: 'Phase 2 accuracy (% of answered recognition trials)', Value: responseRate(recognitionRows) },
      { Section: 'Performance', Measure: 'Phase 2 mean RT on correct trials (ms)', Value: meanRt(correctRecognitionRows) },
      { Section: 'Cueing', Measure: 'Valid-trial mean RT (correct trials only, ms)', Value: validMeanRt },
      { Section: 'Cueing', Measure: 'Invalid-trial mean RT (correct trials only, ms)', Value: invalidMeanRt },
      { Section: 'Cueing', Measure: 'Cueing effect (invalid RT minus valid RT, ms)', Value: validMeanRt === '' || invalidMeanRt === '' ? '' : invalidMeanRt - validMeanRt },
      { Section: 'Cueing', Measure: 'Valid accuracy (correct / all valid trials)', Value: accuracyRate(validPhase1Rows) },
      { Section: 'Cueing', Measure: 'Invalid accuracy (correct / all invalid trials)', Value: accuracyRate(invalidPhase1Rows) },
      { Section: 'Recognition', Measure: 'OLD hit rate (hits / all OLD images)', Value: oldTrials.length ? hitCount / oldTrials.length : '' },
      { Section: 'Recognition', Measure: 'NEW false-alarm rate (false alarms / all NEW images)', Value: newTrials.length ? falseAlarmCount / newTrials.length : '' },
      { Section: 'Recognition', Measure: "d-prime (d') with loglinear correction", Value: oldTrials.length && newTrials.length ? inverseNormalCDF(correctedHitRate) - inverseNormalCDF(correctedFalseAlarmRate) : '' },
      { Section: 'Recognition', Measure: 'Accuracy for OLD valid-cue targets', Value: accuracyRate(oldValidRows) },
      { Section: 'Recognition', Measure: 'Accuracy for OLD invalid-cue targets', Value: accuracyRate(oldInvalidRows) },
      { Section: 'Phase 1', Measure: 'Valid cue trials (cue points to target)', Value: phase1Rows.filter((row) => row.cue_validity === 'valid').length },
      { Section: 'Phase 1', Measure: 'Invalid cue trials (cue points elsewhere)', Value: phase1Rows.filter((row) => row.cue_validity === 'invalid').length },
      { Section: 'Phase 1', Measure: 'Up target trials', Value: phase1Rows.filter((row) => row.match_location === 'up').length },
      { Section: 'Phase 1', Measure: 'Down target trials', Value: phase1Rows.filter((row) => row.match_location === 'down').length },
      { Section: 'Phase 1', Measure: 'Left target trials', Value: phase1Rows.filter((row) => row.match_location === 'left').length },
      { Section: 'Phase 1', Measure: 'Right target trials', Value: phase1Rows.filter((row) => row.match_location === 'right').length },
      { Section: 'Phase 2', Measure: 'Old images (seen during Phase 1)', Value: recognitionRows.filter((row) => row.image_status === 'old').length },
      { Section: 'Phase 2', Measure: 'New images (not previously shown)', Value: recognitionRows.filter((row) => row.image_status === 'new').length },
      { Section: 'Phase 2', Measure: 'Old images correctly called Old (hits)', Value: recognitionRows.filter((row) => row.image_status === 'old' && row.response_meaning === 'Old').length },
      { Section: 'Phase 2', Measure: 'New images incorrectly called Old (false alarms)', Value: recognitionRows.filter((row) => row.image_status === 'new' && row.response_meaning === 'Old').length }
    ];
    const checkRows = [
      ['Phase 1 has 64 trials', phase1Rows.length === 64, phase1Rows.length, 64],
      ['Phase 1 has 48 valid cues', phase1Rows.filter((row) => row.cue_validity === 'valid').length === 48, phase1Rows.filter((row) => row.cue_validity === 'valid').length, 48],
      ['Phase 1 has 16 invalid cues', phase1Rows.filter((row) => row.cue_validity === 'invalid').length === 16, phase1Rows.filter((row) => row.cue_validity === 'invalid').length, 16],
      ['Invalid cues never point to their target', phase1Rows.filter((row) => row.cue_validity === 'invalid' && row.cue_direction === row.match_location).length === 0, phase1Rows.filter((row) => row.cue_validity === 'invalid' && row.cue_direction === row.match_location).length, 0],
      ['Phase 2 has 64 trials', recognitionRows.length === 64, recognitionRows.length, 64],
      ['Phase 2 has 32 old images', recognitionRows.filter((row) => row.image_status === 'old').length === 32, recognitionRows.filter((row) => row.image_status === 'old').length, 32],
      ['Phase 2 has 32 new images', recognitionRows.filter((row) => row.image_status === 'new').length === 32, recognitionRows.filter((row) => row.image_status === 'new').length, 32],
      ['OLD set contains all 16 invalid targets', recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'invalid').length === 16, recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'invalid').length, 16],
      ['OLD set contains 16 valid targets', recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'valid').length === 16, recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'valid').length, 16],
      ...directions.map((direction) => {
        const count = recognitionRows.filter((row) => row.image_status === 'old' && row.source_cue_validity === 'valid' && row.source_target_location === direction).length;
        return [`OLD valid targets balanced at ${direction}`, count === 4, count, 4];
      }),
      ['NEW source Phase 1 fields are blank', recognitionRows.filter((row) => row.image_status === 'new').every((row) => !row.source_phase1_trial && !row.source_cue_validity && !row.source_phase1_response), recognitionRows.filter((row) => row.image_status === 'new' && (row.source_phase1_trial || row.source_cue_validity || row.source_phase1_response)).length, 0],
      ['Phase 2 image IDs are unique', new Set(recognitionRows.map((row) => row.image_id)).size === recognitionRows.length, new Set(recognitionRows.map((row) => row.image_id)).size, recognitionRows.length],
      ['Phase 1 targets are unique', new Set(phase1Rows.map((row) => row.center_image_id)).size === 64, new Set(phase1Rows.map((row) => row.center_image_id)).size, 64],
      ['Phase 1 target/distractor images are not reused', new Set(phase1ImageIds).size === phase1ImageIds.length, new Set(phase1ImageIds).size, phase1ImageIds.length],
      ['Practice images are absent from Phase 1', practiceImageIds.every((id) => !phase1ImageIds.includes(id)), practiceImageIds.filter((id) => phase1ImageIds.includes(id)).length, 0],
      ['All old recognition images appeared in Phase 1', oldRecognitionIds.every((id) => phase1Rows.some((row) => row.center_image_id === id)), oldRecognitionIds.filter((id) => phase1Rows.some((row) => row.center_image_id === id)).length, oldRecognitionIds.length],
      ['No new recognition image appeared in Phase 1 or practice', newRecognitionIds.every((id) => !phase1ImageIds.includes(id) && !practiceImageIds.includes(id)), newRecognitionIds.filter((id) => phase1ImageIds.includes(id) || practiceImageIds.includes(id)).length, 0],
      ...directions.flatMap((direction) => {
        const directionRows = phase1Rows.filter((row) => row.match_location === direction);
        return [
          [`${direction} has 12 valid trials`, directionRows.filter((row) => row.cue_validity === 'valid').length === 12, directionRows.filter((row) => row.cue_validity === 'valid').length, 12],
          [`${direction} has 4 invalid trials`, directionRows.filter((row) => row.cue_validity === 'invalid').length === 4, directionRows.filter((row) => row.cue_validity === 'invalid').length, 4]
        ];
      })
    ].map(([check, passed, observed, expected]) => ({
      Check: check,
      Status: passed ? 'PASS' : 'CHECK',
      Observed: observed,
      Expected: expected
    }));
    const fieldGuideRows = [
      { Column: 'Phase', Meaning: 'Practice, phase1, phase2, or instructions.' },
      { Column: 'Event', Meaning: 'The event recorded on this row: cue, fixation, matching response, recognition response, etc.' },
      { Column: 'Task Code', Meaning: 'The internal event label; filter to phase1_matching or recognition for response analyses.' },
      { Column: 'Cue Validity', Meaning: 'Valid means cue direction equals target location. Invalid means cue direction differs from target location.' },
      { Column: 'Target Location', Meaning: 'Where the matching target appeared; the response key is W=up, A=left, S=down, D=right.' },
      { Column: 'Correct Key / Response Key', Meaning: 'Raw keyboard keys. Matching uses W/A/S/D; recognition uses F=old and J=new.' },
      { Column: 'Response Meaning', Meaning: 'Human-readable interpretation of the participant response key.' },
      { Column: 'Response Outcome', Meaning: 'Correct, Incorrect, No response, or Not applicable (for timed/instruction rows).' },
      { Column: 'Response Time (ms)', Meaning: 'JsPsych response time in milliseconds; blank when no key response was collected.' },
      { Column: 'Memory Status', Meaning: 'Old means shown in Phase 1; New means reserved and not shown before Phase 2.' },
      { Column: 'Accuracy Code', Meaning: 'JsPsych score: 1=correct; 0=incorrect or no response. Use Response Outcome to distinguish omissions from wrong answers. Summary accuracy uses all scheduled trials.' },
      { Column: 'Source Phase 1 fields', Meaning: 'For OLD memory items, identifies their Phase 1 trial, cue condition/location, response, correctness, and RT. Blank for NEW images.' },
      { Column: 'Participant and session fields', Meaning: 'Unique generated participant/session IDs, experiment version and start time, browser, screen, and viewport measurements.' },
      { Column: 'd-prime correction', Meaning: 'Uses the loglinear correction: (hits + 0.5)/(OLD count + 1) and (false alarms + 0.5)/(NEW count + 1), preventing infinite d-prime.' }
    ];
    const workbook = XLSX.utils.book_new();
    const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
    const resultsSheet = XLSX.utils.json_to_sheet(resultRows, { header: exportColumns.map(([label]) => label) });
    const checksSheet = XLSX.utils.json_to_sheet(checkRows);
    const guideSheet = XLSX.utils.json_to_sheet(fieldGuideRows);
    summaryRows.forEach((row, index) => {
      if (row.Measure.includes('accuracy (%') || row.Measure.includes('accuracy (correct /') || row.Measure.includes(' rate (') || row.Measure.includes('Accuracy for OLD')) {
        const valueCell = summarySheet[XLSX.utils.encode_cell({ r: index + 1, c: 2 })];
        if (valueCell && valueCell.v !== '') valueCell.z = '0.0%';
      }
    });
    summarySheet['!cols'] = [{ wch: 18 }, { wch: 54 }, { wch: 16 }];
    resultsSheet['!cols'] = exportColumns.map(([label]) => ({ wch: Math.min(Math.max(label.length + 2, 16), 32) }));
    resultsSheet['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(exportColumns.length - 1)}${resultRows.length + 1}` };
    checksSheet['!cols'] = [{ wch: 34 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
    guideSheet['!cols'] = [{ wch: 32 }, { wch: 100 }];
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
    XLSX.utils.book_append_sheet(workbook, resultsSheet, 'Results');
    XLSX.utils.book_append_sheet(workbook, checksSheet, 'Data Checks');
    XLSX.utils.book_append_sheet(workbook, guideSheet, 'Field Guide');
    const sessionSheet = XLSX.utils.json_to_sheet(Object.entries(sessionMetadata).map(([field, value]) => ({
      Field: field,
      Value: value
    })));
    sessionSheet['!cols'] = [{ wch: 34 }, { wch: 100 }];
    XLSX.utils.book_append_sheet(workbook, sessionSheet, 'Session Info');
    const workbookData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const timestamp = new Date().toISOString()
      .replace('T', '_')
      .replace(/:/g, '')
      .slice(0, 16);
    const filename = `posner_memory_${timestamp}.xlsx`;
    const downloadUrl = URL.createObjectURL(new Blob([workbookData], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
    const download = () => {
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
    };

    void saveDataRemotely(JSON.stringify(rawTrials));
    download();
    document.getElementById('jspsych-target').innerHTML = `
      <section class="instruction">
        <h1>Finished</h1>
        <p>Thank you. Your responses have been recorded.</p>
        <button class="jspsych-btn" id="download-results">Download Results</button>
      </section>`;
    document.getElementById('download-results').addEventListener('click', download);
    setTimeout(() => URL.revokeObjectURL(downloadUrl), 60000);
  }
});

jsPsych.data.addProperties(sessionMetadata);

async function saveDataRemotely(dataJson) {
  if (!REMOTE_DATA_CONFIG.enabled) return { status: 'disabled' };
  if (!REMOTE_DATA_CONFIG.experimentId) {
    console.warn('Remote data saving is enabled but its DataPipe experiment ID is not configured.');
    return { status: 'not_configured' };
  }
  console.info('Remote data upload integration is not configured yet.', {
    provider: REMOTE_DATA_CONFIG.provider,
    experimentId: REMOTE_DATA_CONFIG.experimentId,
    dataBytes: dataJson.length
  });
  return { status: 'integration_required' };
}

function describeResponse(response, task) {
  if (!response) return 'No response';
  if (task === 'recognition') {
    return { f: 'Old', j: 'New' }[response.toLowerCase()] ?? `Other key: ${response}`;
  }
  if (task === 'practice_matching' || task === 'phase1_matching') {
    return { w: 'Up', a: 'Left', s: 'Down', d: 'Right' }[response.toLowerCase()] ?? `Other key: ${response}`;
  }
  return 'Instruction keypress';
}

function inverseNormalCDF(probability) {
  const a = [-39.6968302866538, 220.946098424521, -275.928510446969, 138.357751867269, -30.6647980661472, 2.50662827745924];
  const b = [-54.4760987982241, 161.585836858041, -155.698979859887, 66.8013118877197, -13.2806815528857];
  const c = [-0.00778489400243029, -0.322396458041136, -2.40075827716184, -2.54973253934373, 4.37466414146497, 2.93816398269878];
  const d = [0.00778469570904146, 0.32246712907004, 2.445134137143, 3.75440866190742];
  const tailThreshold = 0.02425;

  if (probability <= 0 || probability >= 1) return Number.NaN;
  if (probability < tailThreshold) {
    const q = Math.sqrt(-2 * Math.log(probability));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  if (probability > 1 - tailThreshold) {
    const q = Math.sqrt(-2 * Math.log(1 - probability));
    return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }

  const centered = probability - 0.5;
  const squared = centered * centered;
  return (((((a[0] * squared + a[1]) * squared + a[2]) * squared + a[3]) * squared + a[4]) * squared + a[5]) * centered /
    (((((b[0] * squared + b[1]) * squared + b[2]) * squared + b[3]) * squared + b[4]) * squared + 1);
}

const objectDatasetDirectory = 'stimuli/OBJECTSALL/';
const objectManifestPath = 'stimuli/object_manifest.json';
const imagePath = (path) => path;

async function discoverObjectImages() {
  const response = await fetch(objectManifestPath);

  if (!response.ok) {
    throw new Error(`Failed to load stimulus manifest (${response.status}).`);
  }

  const paths = await response.json();

  const uniquePaths = [...new Set(paths)]
    .filter((path) => typeof path === 'string')
    .filter((path) => path.startsWith(objectDatasetDirectory))
    .filter((path) => /\.(?:jpg|jpeg|png|webp)$/i.test(path));

  if (uniquePaths.length < 292) {
    throw new Error(`At least 292 Brady images are required; found ${uniquePaths.length}.`);
  }

  return uniquePaths;
}

const directions = ['up', 'left', 'down', 'right'];
const keysByDirection = { up: 'w', left: 'a', down: 's', right: 'd' };
const arrowsByDirection = { up: '↑', left: '←', down: '↓', right: '→' };
const gridPositionByDirection = { up: 1, left: 3, down: 7, right: 5 };
const responseKeys = ['w', 'a', 's', 'd'];

const shuffled = (items) => jsPsych.randomization.shuffle([...items]);

function makePhase1Trial(matchLocation, valid, centerImageId, distractorIds, trialIndex, isPractice = false) {
  const matchImageId = centerImageId;
  const cueDirection = valid
    ? matchLocation
    : shuffled(directions.filter((direction) => direction !== matchLocation))[0];
  const imagesByPosition = Array(9).fill(null);
  imagesByPosition[4] = centerImageId;
  imagesByPosition[gridPositionByDirection[matchLocation]] = matchImageId;
  directions.filter((direction) => direction !== matchLocation).forEach((direction, index) => {
    imagesByPosition[gridPositionByDirection[direction]] = distractorIds[index];
  });

  return {
    type: htmlKeyboardResponse,
    stimulus: `<div class="match-grid">${imagesByPosition.map((id, index) => id
      ? `<div class="image-slot"><img src="${imagePath(id)}" alt=""></div>`
      : `<div class="image-slot empty" aria-hidden="true"></div>`).join('')}</div>`,
    choices: responseKeys,
    data: {
      task: isPractice ? 'practice_matching' : 'phase1_matching',
      phase: isPractice ? 'practice' : 'phase1',
      trial_type: 'visual_matching',
      trial_index: trialIndex,
      center_image_id: centerImageId,
      match_image_id: matchImageId,
      distractor_image_ids: distractorIds.join(','),
      cue_direction: cueDirection,
      validity: valid ? 'valid' : 'invalid',
      match_location: matchLocation,
      correct_response: keysByDirection[matchLocation]
    },
    on_finish: (data) => {
      data.response = data.response ?? '';
      data.accuracy = data.response === data.correct_response ? 1 : 0;
    }
  };
}

function makeRecognitionTrial(imageId, trialIndex, sourceTrial = null) {
  const isOld = sourceTrial !== null;
  return {
    type: htmlKeyboardResponse,
    stimulus: `<img class="memory-image" src="${imagePath(imageId)}" alt="">\n<div class="response-prompt">F = OLD &nbsp;&nbsp;&nbsp; J = NEW</div>`,
    choices: ['f', 'j'],
    data: {
      task: 'recognition',
      phase: 'phase2',
      trial_type: 'surprise_recognition',
      trial_index: trialIndex,
      image_id: imageId,
      image_status: isOld ? 'old' : 'new',
      correct_response: isOld ? 'f' : 'j',
      source_phase1_trial: sourceTrial?.data.trial_index ?? null,
      source_cue_validity: sourceTrial?.data.validity ?? null,
      source_target_location: sourceTrial?.data.match_location ?? null,
      source_cue_direction: sourceTrial?.data.cue_direction ?? null,
      source_phase1_response: null,
      source_phase1_correct: null,
      source_phase1_rt: null
    },
    on_finish: (data) => {
      data.response = data.response ?? '';
      data.accuracy = data.response === data.correct_response ? 1 : 0;
      if (isOld) {
        const sourceData = jsPsych.data.get().filter({ task: 'phase1_matching' }).values()
          .find((row) => row.center_image_id === imageId);
        data.source_phase1_response = sourceData?.participant_response ?? null;
        data.source_phase1_correct = sourceData?.accuracy ?? null;
        data.source_phase1_rt = sourceData?.rt ?? null;
      }
    }
  };
}

const instructions = (title, body) => ({
  type: htmlKeyboardResponse,
  stimulus: `<section class="instruction"><h1>${title}</h1><p>${body}</p><p>Press any key to continue.</p></section>`,
  choices: 'ALL_KEYS',
  data: { task: 'instructions', phase: 'instructions', trial_type: 'instructions' }
});

const datasetFiles = await discoverObjectImages();
const allocatedImages = shuffled(datasetFiles);
const practiceImageIds = allocatedImages.slice(0, 4);
const phaseImageIds = allocatedImages.slice(4, 68);
const phaseDistractorIds = allocatedImages.slice(68, 260);
const newImageIds = allocatedImages.slice(260, 292);
const preloadImagePaths = [...practiceImageIds, ...phaseImageIds, ...phaseDistractorIds, ...newImageIds];

const practiceTrials = directions.map((matchLocation, index) =>
  makePhase1Trial(
    matchLocation,
    true,
    practiceImageIds[index],
    practiceImageIds.filter((imageId) => imageId !== practiceImageIds[index]),
    index + 1,
    true
  )
);
const practiceTimeline = practiceTrials.flatMap((trial) => [
  { type: htmlKeyboardResponse, stimulus: '<div class="fixation">+</div>', trial_duration: 500, choices: 'NO_KEYS', data: { task: 'practice_fixation', phase: 'practice', trial_type: 'fixation', duration_ms: 500 } },
  { type: htmlKeyboardResponse, stimulus: `<div class="cue">${arrowsByDirection[trial.data.cue_direction]}</div>`, trial_duration: 200, choices: 'NO_KEYS', data: { task: 'practice_cue', phase: 'practice', trial_type: 'cue', cue_direction: trial.data.cue_direction, duration_ms: 200 } },
  { type: htmlKeyboardResponse, stimulus: '<div class="fixation">+</div>', trial_duration: 300, choices: 'NO_KEYS', data: { task: 'practice_postcue_fixation', phase: 'practice', trial_type: 'postcue_fixation', duration_ms: 300 } },
  trial,
  {
    type: htmlKeyboardResponse,
    stimulus: () => jsPsych.data.get().last(1).values()[0]?.accuracy === 1 ? '<p class="instruction">Correct</p>' : '<p class="instruction">Incorrect</p>',
    trial_duration: 500,
    choices: 'NO_KEYS',
    data: { task: 'practice_feedback', phase: 'practice', trial_type: 'feedback', duration_ms: 500 }
  }
]);
const practiceLoop = {
  timeline: practiceTimeline,
  loop_function: (data) => data.filter({ task: 'practice_matching' }).values()
    .filter((trial) => trial.accuracy === 1).length < 3
};
const phase1Schedule = shuffled(directions.flatMap((matchLocation) => [
  ...Array(12).fill({ matchLocation, valid: true }),
  ...Array(4).fill({ matchLocation, valid: false })
]));
const phase1Trials = shuffled(phaseImageIds).map((centerImageId, index) =>
  makePhase1Trial(
    phase1Schedule[index].matchLocation,
    phase1Schedule[index].valid,
    centerImageId,
    phaseDistractorIds.slice(index * 3, index * 3 + 3),
    index + 1
  )
);
const invalidOldSources = phase1Trials.filter((trial) => trial.data.validity === 'invalid');
const validOldSourcesByLocation = directions.map((direction) => shuffled(
  phase1Trials.filter((trial) => trial.data.validity === 'valid' && trial.data.match_location === direction)
).slice(0, 4));
const recognitionOldSources = [
  ...invalidOldSources,
  ...validOldSourcesByLocation.flat()
];
const recognitionOldIds = recognitionOldSources.map((trial) => trial.data.center_image_id);
const recognitionTrials = shuffled([
  ...recognitionOldSources.map((sourceTrial) => ({
    imageId: sourceTrial.data.center_image_id,
    sourceTrial
  })),
  ...newImageIds.map((imageId) => ({ imageId, sourceTrial: null }))
]).map(({ imageId, sourceTrial }, index) => makeRecognitionTrial(imageId, index + 1, sourceTrial));

const timeline = [
  { type: preloadPlugin, images: preloadImagePaths, message: '<p>Loading experiment...</p>', show_progress_bar: true, continue_after_error: false, show_detailed_errors: true },
  instructions('Visual attention task', 'Keep your eyes on the center of the screen. Each trial begins with a fixation cross and an arrow, followed by a group of images. The arrow indicates where the matching image is likely to appear, but it is not always correct. Respond with W for up, A for left, S for down, or D for right, according to the image that matches the center image.'),
  instructions('Practice', 'You will first complete a short practice block.'),
  practiceLoop,
  instructions('Main task', 'The main task begins now. Respond as quickly and accurately as you can. There will be no feedback during the task.'),
  ...phase1Trials.flatMap((trial) => [
    { type: htmlKeyboardResponse, stimulus: '<div class="fixation">+</div>', trial_duration: 500, choices: 'NO_KEYS', data: { task: 'phase1_fixation', phase: 'phase1', trial_type: 'fixation', duration_ms: 500 } },
    { type: htmlKeyboardResponse, stimulus: `<div class="cue">${arrowsByDirection[trial.data.cue_direction]}</div>`, trial_duration: 200, choices: 'NO_KEYS', data: { task: 'phase1_cue', phase: 'phase1', trial_type: 'cue', cue_direction: trial.data.cue_direction, duration_ms: 200 } },
    { type: htmlKeyboardResponse, stimulus: '<div class="fixation">+</div>', trial_duration: 300, choices: 'NO_KEYS', data: { task: 'phase1_postcue_fixation', phase: 'phase1', trial_type: 'postcue_fixation', duration_ms: 300 } },
    trial
  ]),
  instructions('Memory test', 'You will now be tested on your memory for the images. For each image, press F if you saw it earlier, or J if it is new.'),
  ...recognitionTrials,
  instructions('Finished', 'Thank you. Your responses have been recorded.')
];

const referencedStimulusPaths = timeline.flatMap((trial) => {
  if (typeof trial.stimulus !== 'string') return [];
  return [...trial.stimulus.matchAll(/src="([^"]+)"/g)].map((match) => match[1]);
});

async function validateStimulusPaths(paths) {
  const errors = [];
  const malformedPaths = paths.filter((path) => !/^stimuli\/.+\.(?:svg|png|jpe?g|webp)$/i.test(path));
  const undefinedPaths = paths.filter((path) => path.includes('undefined') || path.includes('null'));
  const duplicatePreloadPaths = preloadImagePaths.filter((path, index) => preloadImagePaths.indexOf(path) !== index);
  const recognitionImageIds = recognitionTrials.map((trial) => trial.data.image_id);
  const duplicateRecognitionIds = recognitionImageIds.filter((id, index) => recognitionImageIds.indexOf(id) !== index);
  const phaseTargetIds = phase1Trials.map((trial) => trial.data.center_image_id);
  const duplicatePhaseTargets = phaseTargetIds.filter((id, index) => phaseTargetIds.indexOf(id) !== index);
  const phaseCounts = directions.map((direction) => ({
    direction,
    total: phase1Trials.filter((trial) => trial.data.match_location === direction).length,
    valid: phase1Trials.filter((trial) => trial.data.match_location === direction && trial.data.validity === 'valid').length,
    invalid: phase1Trials.filter((trial) => trial.data.match_location === direction && trial.data.validity === 'invalid').length
  }));
  const invalidCueMatches = phase1Trials.filter((trial) => trial.data.validity === 'invalid'
    && trial.data.cue_direction === trial.data.match_location);
  const malformedGridTrials = [...practiceTrials, ...phase1Trials].filter((trial) => {
    const distractors = trial.data.distractor_image_ids.split(',').filter(Boolean);
    return distractors.length !== 3 || new Set([trial.data.center_image_id, ...distractors]).size !== 4;
  });

  if (malformedPaths.length > 0) errors.push(`Malformed stimulus paths: ${malformedPaths.join(', ')}`);
  if (undefinedPaths.length > 0) errors.push(`Undefined stimulus paths: ${undefinedPaths.join(', ')}`);
  if (duplicatePreloadPaths.length > 0) errors.push(`Duplicate preload entries: ${[...new Set(duplicatePreloadPaths)].join(', ')}`);
  if (duplicateRecognitionIds.length > 0) errors.push(`Duplicate recognition images: ${[...new Set(duplicateRecognitionIds)].join(', ')}`);
  if (phase1Trials.length !== 64) errors.push(`Phase 1 requires 64 trials, found ${phase1Trials.length}.`);
  if (invalidCueMatches.length > 0) errors.push(`Invalid cue points to the matching location on ${invalidCueMatches.length} trials.`);
  if (malformedGridTrials.length > 0) errors.push(`${malformedGridTrials.length} matching trials do not contain a target and three unique distractors.`);
  if (duplicatePhaseTargets.length > 0) errors.push(`Phase 1 target images are reused: ${[...new Set(duplicatePhaseTargets)].join(', ')}`);
  phaseCounts.forEach(({ direction, total, valid, invalid }) => {
    if (total !== 16 || valid !== 12 || invalid !== 4) {
      errors.push(`Unbalanced Phase 1 ${direction} trials: ${total} total, ${valid} valid, ${invalid} invalid.`);
    }
  });
  if (recognitionOldIds.length !== newImageIds.length) errors.push('Recognition old/new sets are not equal in size.');
  if (recognitionOldSources.length !== 32) errors.push(`Recognition requires 32 OLD images, found ${recognitionOldSources.length}.`);
  if (invalidOldSources.length !== 16) errors.push(`Recognition requires all 16 invalid-target images, found ${invalidOldSources.length}.`);
  if (validOldSourcesByLocation.some((sources) => sources.length !== 4)) errors.push('OLD valid targets must include four trials from each target location.');
  if (newImageIds.some((id) => phaseImageIds.includes(id) || phaseDistractorIds.includes(id) || practiceImageIds.includes(id))) {
    errors.push('Recognition NEW images overlap with practice or Phase 1 images.');
  }
  if (new Set([...practiceImageIds, ...phaseImageIds, ...phaseDistractorIds, ...newImageIds]).size !== 292) {
    errors.push('Practice, Phase 1, and Phase 2 image allocations must contain 292 unique images.');
  }

  if (errors.length > 0) {
    console.error('Stimulus validation failed:', errors);
    throw new Error('Stimulus validation failed. See the console for details.');
  }
}

const allStimulusPaths = [...preloadImagePaths, ...referencedStimulusPaths];
try {
  await validateStimulusPaths(allStimulusPaths);
  console.log(`Validated ${preloadImagePaths.length} unique selected stimulus images; jsPsych preload will verify and load them.`);
} catch (error) {
  console.error(error);
  document.getElementById('jspsych-target').innerHTML = '<section class="instruction"><h1>Experiment setup error</h1><p>Stimulus validation failed. Check the browser console for details.</p></section>';
  throw error;
}
document.getElementById('jspsych-target').innerHTML = '';
jsPsych.run(timeline);
