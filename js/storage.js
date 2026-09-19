const STORE_KEY = 'candeApp_v1';

function todayISO() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

function loadStore() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return defaultStore();
    const parsed = JSON.parse(raw);
    return { ...defaultStore(), ...parsed };
  } catch (e) {
    return defaultStore();
  }
}

function defaultStore() {
  return {
    settings: { startDate: todayISO() },
    foodLogs: {},
    workoutLogs: {},
    habits: {},
    weightLog: [],
    measurementLog: []
  };
}

function saveStore(store) {
  localStorage.setItem(STORE_KEY, JSON.stringify(store));
}

let STORE = loadStore();

function persist() {
  saveStore(STORE);
}

function getFoodLog(dateISO) {
  return STORE.foodLogs[dateISO] || [];
}

function addFoodEntry(dateISO, entry) {
  if (!STORE.foodLogs[dateISO]) STORE.foodLogs[dateISO] = [];
  STORE.foodLogs[dateISO].push({ id: Date.now() + Math.random().toString(16).slice(2), ...entry });
  persist();
}

function removeFoodEntry(dateISO, entryId) {
  if (!STORE.foodLogs[dateISO]) return;
  STORE.foodLogs[dateISO] = STORE.foodLogs[dateISO].filter(e => e.id !== entryId);
  persist();
}

function getWorkoutLog(dateISO) {
  return STORE.workoutLogs[dateISO] || { workoutKey: null, completed: [] };
}

function toggleExerciseDone(dateISO, workoutKey, exerciseIndex) {
  if (!STORE.workoutLogs[dateISO]) STORE.workoutLogs[dateISO] = { workoutKey, completed: [] };
  const log = STORE.workoutLogs[dateISO];
  log.workoutKey = workoutKey;
  const pos = log.completed.indexOf(exerciseIndex);
  if (pos >= 0) log.completed.splice(pos, 1);
  else log.completed.push(exerciseIndex);
  persist();
}

function getHabits(dateISO) {
  return STORE.habits[dateISO] || { walk: false, water: false, dayDone: false };
}

function toggleHabit(dateISO, key) {
  if (!STORE.habits[dateISO]) STORE.habits[dateISO] = { walk: false, water: false, dayDone: false };
  STORE.habits[dateISO][key] = !STORE.habits[dateISO][key];
  persist();
}

function addWeightEntry(dateISO, weightKg) {
  STORE.weightLog = STORE.weightLog.filter(e => e.date !== dateISO);
  STORE.weightLog.push({ date: dateISO, weight: weightKg });
  STORE.weightLog.sort((a, b) => a.date.localeCompare(b.date));
  persist();
}

function removeWeightEntry(dateISO) {
  STORE.weightLog = STORE.weightLog.filter(e => e.date !== dateISO);
  persist();
}

function addMeasurementEntry(dateISO, data) {
  STORE.measurementLog = STORE.measurementLog.filter(e => e.date !== dateISO);
  STORE.measurementLog.push({ date: dateISO, ...data });
  STORE.measurementLog.sort((a, b) => a.date.localeCompare(b.date));
  persist();
}

function removeMeasurementEntry(dateISO) {
  STORE.measurementLog = STORE.measurementLog.filter(e => e.date !== dateISO);
  persist();
}

function setStartDate(dateISO) {
  STORE.settings.startDate = dateISO;
  persist();
}

function currentWeekNumber() {
  const start = new Date(STORE.settings.startDate + 'T00:00:00');
  const now = new Date(todayISO() + 'T00:00:00');
  const diffDays = Math.floor((now - start) / (1000 * 60 * 60 * 24));
  const week = Math.floor(diffDays / 7) + 1;
  return Math.min(Math.max(week, 1), 8);
}

function computeStreak() {
  let streak = 0;
  let d = new Date(todayISO() + 'T00:00:00');
  for (let i = 0; i < 365; i++) {
    const iso = d.toISOString().slice(0, 10);
    const h = STORE.habits[iso];
    if (h && h.dayDone) {
      streak++;
      d.setDate(d.getDate() - 1);
    } else {
      break;
    }
  }
  return streak;
}
