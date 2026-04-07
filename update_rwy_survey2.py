#!/usr/bin/env python3
import re

file_path = '/Users/jacobanderson/mission-briefing-app/PilotApp.html'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update 1: Update all prompts state initializations to include c50 fields
# First, in the initial statusGPS display function update
old_status = '''    const el = document.getElementById('rwysurvey-status');
    if (el) {
      const gps = runwaySurveyToolState.gps || {};
      let txt = gps.tracking ? 'GPS active' : 'GPS idle';
      txt += ' · Acc: ±' + Math.round(Number(gps.current.acc || 0)) + 'm';
      el.textContent = txt;
    }'''

new_status = '''    const el = document.getElementById('rwysurvey-status');
    if (el) {
      const gps = runwaySurveyToolState.gps || {};
      let txt = gps.tracking ? 'GPS active' : 'GPS idle';
      txt += ' · Acc: ±' + Math.round(Number(gps.current.acc || 0)) + 'm';
      const marks = Array.isArray(runwaySurveyToolState.perimeter.cornerMarks) ? runwaySurveyToolState.perimeter.cornerMarks.length : 0;
      if (marks === 0 && gps.tracking) txt += ' · (GPS warming up — stand still at corner before marking)';
      el.textContent = txt;
    }'''

content = content.replace(old_status, new_status)

# Update 2: Update prompts state in _rwySurveyOnRunwayChange_ function
old_on_rwy = '''    runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
    runwaySurveyToolState.ui = { pausedByPopup: false };
    runwaySurveyToolState.prompts = { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null };
    runwaySurveyToolState.widthObservations = [];'''

new_on_rwy = '''    runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
    runwaySurveyToolState.ui = { pausedByPopup: false };
    runwaySurveyToolState.prompts = { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null, c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null };
    runwaySurveyToolState.widthObservations = [];'''

content = content.replace(old_on_rwy, new_on_rwy)

# Update 3: Fix any remaining single prompts initialization in clearRunwaySurveyTrace
old_clear = '''    runwaySurveyToolState.obstacleAngles50m = [];
    runwaySurveyToolState.prompts = { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null };
    runwaySurveyToolState.slopeSegments = [];'''

new_clear = '''    runwaySurveyToolState.obstacleAngles50m = [];
    runwaySurveyToolState.prompts = { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null, c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null };
    runwaySurveyToolState.slopeSegments = [];'''

if old_clear in content:
    content = content.replace(old_clear, new_clear)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Updated prompts state to include c50 checkpoint fields")
print("✓ Added GPS warm-up message to status display")
print("\nAll updates completed successfully!")
