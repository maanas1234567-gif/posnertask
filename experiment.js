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
      ['Trial Number', 'trial_number'],
      ['Phase', 'phase'],
      ['Trial Type', 'trial_type'],
      ['Task', 'task'],
      ['Trial Within Phase', 'trial_index'],
      ['Recognition Image', 'image_id'],
      ['Center Image', 'center_image_id'],
      ['Matching Image', 'match_image_id'],
      ['Distractor Images', 'distractor_image_ids'],
      ['Cue Direction', 'cue_direction'],
      ['Cue Validity', 'cue_validity'],
      ['Match Location', 'match_location'],
      ['Correct Response', 'correct_response'],
      ['Participant Response', 'participant_response'],
      ['Accuracy (1=correct, 0=incorrect)', 'accuracy'],
      ['Response Time (ms)', 'rt_ms'],
      ['Duration (ms)', 'duration_ms'],
      ['Memory Status', 'image_status']
    ];
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
      participant_response: trial.participant_response || 'no_response',
      accuracy: trial.accuracy ?? '',
      rt_ms: trial.rt ?? '',
      duration_ms: trial.duration_ms ?? '',
      image_status: trial.image_status ?? ''
    }));
    const resultRows = exportRows.map((row) => Object.fromEntries(
      exportColumns.map(([label, key]) => [label, row[key]])
    ));
    const summaryRows = [
      { Measure: 'Total recorded rows', Value: exportRows.length },
      { Measure: 'Practice matching trials', Value: exportRows.filter((row) => row.task === 'practice_matching').length },
      { Measure: 'Phase 1 matching trials', Value: exportRows.filter((row) => row.task === 'phase1_matching').length },
      { Measure: 'Phase 2 recognition trials', Value: exportRows.filter((row) => row.task === 'recognition').length },
      { Measure: 'Matching and recognition responses', Value: exportRows.filter((row) => row.participant_response !== 'no_response').length },
      { Measure: 'Correct responses', Value: exportRows.filter((row) => row.accuracy === 1).length },
      { Measure: 'Incorrect responses', Value: exportRows.filter((row) => row.accuracy === 0).length }
    ];
    const workbook = XLSX.utils.book_new();
    const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
    const resultsSheet = XLSX.utils.json_to_sheet(resultRows, { header: exportColumns.map(([label]) => label) });
    summarySheet['!cols'] = [{ wch: 38 }, { wch: 18 }];
    resultsSheet['!cols'] = exportColumns.map(([label]) => ({ wch: Math.min(Math.max(label.length + 2, 16), 32) }));
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
    XLSX.utils.book_append_sheet(workbook, resultsSheet, 'Results');
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
