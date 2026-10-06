import { initJsPsych } from 'https://cdn.jsdelivr.net/npm/jspsych@8.2.1/+esm';
import htmlKeyboardResponse from 'https://cdn.jsdelivr.net/npm/@jspsych/plugin-html-keyboard-response@2.1.0/+esm';
import * as XLSX from 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';

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
      ['Accuracy Code', 'accuracy']
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
      image_status: trial.image_status ?? ''
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
      ['Phase 2 has 64 trials', recognitionRows.length === 64, recognitionRows.length, 64],
      ['Phase 2 has 32 old images', recognitionRows.filter((row) => row.image_status === 'old').length === 32, recognitionRows.filter((row) => row.image_status === 'old').length, 32],
      ['Phase 2 has 32 new images', recognitionRows.filter((row) => row.image_status === 'new').length === 32, recognitionRows.filter((row) => row.image_status === 'new').length, 32],
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
      { Column: 'Accuracy Code', Meaning: 'JsPsych score: 1=correct; 0=incorrect or no response. Use Response Outcome to distinguish omissions from wrong answers.' }
    ];
    const workbook = XLSX.utils.book_new();
    const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
    const resultsSheet = XLSX.utils.json_to_sheet(resultRows, { header: exportColumns.map(([label]) => label) });
    const checksSheet = XLSX.utils.json_to_sheet(checkRows);
    const guideSheet = XLSX.utils.json_to_sheet(fieldGuideRows);
    summaryRows.forEach((row, index) => {
      if (row.Measure.includes('accuracy (%')) {
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

    download();
    document.getElementById('jspsych-target').innerHTML = `
      <section class="instruction">
        <h1>Finished</h1>
        <p>Thank you. Your responses have been recorded.</p>
        <button class="jspsych-btn" id="download-results">Download Results Excel File</button>
      </section>`;
    document.getElementById('download-results').addEventListener('click', download);
    setTimeout(() => URL.revokeObjectURL(downloadUrl), 60000);
  }
});

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

function makeRecognitionTrial(imageId, trialIndex) {
  const isOld = phaseImageIds.includes(imageId);
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
      correct_response: isOld ? 'f' : 'j'
    },
    on_finish: (data) => {
      data.response = data.response ?? '';
      data.accuracy = data.response === data.correct_response ? 1 : 0;
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
  trial
]);
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
const recognitionOldIds = shuffled(phaseImageIds).slice(0, newImageIds.length);
const recognitionTrials = shuffled([...recognitionOldIds, ...newImageIds])
  .map((imageId, index) => makeRecognitionTrial(imageId, index + 1));

const timeline = [
  instructions('Visual attention task', 'Keep your eyes on the center of the screen. Each trial begins with a fixation cross and an arrow, followed by a group of images. Respond with W for up, A for left, S for down, or D for right, according to the image that matches the center image.'),
  instructions('Practice', 'You will first complete a short practice block.'),
  ...practiceTimeline,
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

  if (malformedPaths.length > 0) errors.push(`Malformed stimulus paths: ${malformedPaths.join(', ')}`);
  if (undefinedPaths.length > 0) errors.push(`Undefined stimulus paths: ${undefinedPaths.join(', ')}`);
  if (duplicatePreloadPaths.length > 0) errors.push(`Duplicate preload entries: ${[...new Set(duplicatePreloadPaths)].join(', ')}`);
  if (duplicateRecognitionIds.length > 0) errors.push(`Duplicate recognition images: ${[...new Set(duplicateRecognitionIds)].join(', ')}`);
  if (phase1Trials.length !== 64) errors.push(`Phase 1 requires 64 trials, found ${phase1Trials.length}.`);
  if (duplicatePhaseTargets.length > 0) errors.push(`Phase 1 target images are reused: ${[...new Set(duplicatePhaseTargets)].join(', ')}`);
  phaseCounts.forEach(({ direction, total, valid, invalid }) => {
    if (total !== 16 || valid !== 12 || invalid !== 4) {
      errors.push(`Unbalanced Phase 1 ${direction} trials: ${total} total, ${valid} valid, ${invalid} invalid.`);
    }
  });
  if (recognitionOldIds.length !== newImageIds.length) errors.push('Recognition old/new sets are not equal in size.');
  if (newImageIds.some((id) => phaseImageIds.includes(id) || practiceImageIds.includes(id))) {
    errors.push('Recognition new images overlap with Phase 1 or practice images.');
  }

  const checkedPaths = [...new Set(paths)];
  await Promise.all(checkedPaths.map(async (path) => {
    try {
      const response = await fetch(path);
      if (!response.ok) errors.push(`Missing stimulus (${response.status}): ${path}`);
      else await response.blob();
    } catch (error) {
      errors.push(`Failed to request stimulus: ${path} (${error.message})`);
    }
  }));

  if (errors.length > 0) {
    console.error('Stimulus validation failed:', errors);
    throw new Error('Stimulus validation failed. See the console for details.');
  }
}

async function preloadImages(paths) {
  for (const path of paths) {
    const response = await fetch(path);
    if (!response.ok) throw new Error(`Failed to preload stimulus (${response.status}): ${path}`);
    await response.blob();
  }
}

const allStimulusPaths = [...preloadImagePaths, ...referencedStimulusPaths];
await validateStimulusPaths(allStimulusPaths);
await preloadImages(preloadImagePaths);
console.log(`Validated and preloaded ${preloadImagePaths.length} stimulus images.`);
jsPsych.run(timeline);
