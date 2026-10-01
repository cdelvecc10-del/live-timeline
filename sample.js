/*
 * Live Timeline — built-in example (fictional), used when the pane is opened outside PowerPoint
 * and for "Start with example scenarios" on a slide that has a timeline table.
 */
(function (root) {
  'use strict';
  var E = root.LT.engine;
  var names = ['A. Immunobridging study', 'B. Efficacy study', 'C. File on existing data'];
  root.LT.sample = {
    title: 'Example: Vaccine X US filing scenarios',
    names: names,
    baselineLabel: 'Example baseline',
    model: {
      title: 'Example: Vaccine X US filing scenarios',
      rules: { reviewPriority: 8, reviewStandard: 12, acipMonths: [2, 6, 10], acipMinGap: 1, readoutLag: 2, submissionLag: 1, pointLabels: 'quarter' },
      scenarios: [
        { id: 's1', name: names[0], row: 0, studyName: 'Ph3 IB', ssd: '2027-03', studyMonths: 9, priorityReview: false },
        { id: 's2', name: names[1], row: 1, studyName: 'Ph3 efficacy', ssd: '2027-10', studyMonths: 14, priorityReview: false },
        { id: 's3', name: names[2], row: 2, studyName: '', ssd: '', studyMonths: '', submission: '2027-02', priorityReview: true, note: 'Resubmission with updated analysis' }
      ]
    },
    cal: {
      axis: { startMonth: E.mi(2026, 1), endMonth: E.mi(2031, 1), left: 179, right: 697 },
      lanes: [{ top: 146, bottom: 250 }, { top: 250, bottom: 354 }, { top: 354, bottom: 458 }],
      gridTop: 146, gridBottom: 458, source: 'example'
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
