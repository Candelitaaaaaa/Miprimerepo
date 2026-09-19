let currentTab = 'hoy';
let selectedWorkoutKey = null;

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  Object.entries(attrs).forEach(([k, v]) => {
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  });
  (Array.isArray(children) ? children : [children]).forEach(c => {
    if (c === null || c === undefined) return;
    node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  });
  return node;
}

function fmt1(n) {
  return Math.round(n * 10) / 10;
}

function switchTab(tab) {
  currentTab = tab;
  document.querySelectorAll('.navbtn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  render();
}

function render() {
  const root = document.getElementById('view');
  root.innerHTML = '';
  if (currentTab === 'hoy') root.appendChild(renderHoy());
  if (currentTab === 'entreno') root.appendChild(renderEntreno());
  if (currentTab === 'comidas') root.appendChild(renderComidas());
  if (currentTab === 'progreso') root.appendChild(renderProgreso());
  if (currentTab === 'plan') root.appendChild(renderPlan());
}

function card(children, extraClass = '') {
  return el('div', { class: 'card ' + extraClass }, children);
}

function progressBar(value, max, colorClass) {
  const pct = Math.min(100, Math.round((value / max) * 100));
  return el('div', { class: 'bar-track' }, el('div', { class: 'bar-fill ' + colorClass, style: `width:${pct}%` }));
}

/* ---------- HOY ---------- */

function renderHoy() {
  const today = todayISO();
  const jsDay = new Date(today + 'T00:00:00').getDay();
  const dayName = DAY_NAMES[jsDay];
  const schedule = WEEKLY_SCHEDULE.find(s => s.day === dayName);
  const week = currentWeekNumber();
  const phase = phaseForWeek(week);
  const foods = getFoodLog(today);
  const kcal = foods.reduce((s, f) => s + f.kcal, 0);
  const prot = foods.reduce((s, f) => s + f.prot, 0);
  const habits = getHabits(today);
  const streak = computeStreak();
  const wlog = getWorkoutLog(today);

  const frag = el('div', {});

  frag.appendChild(el('div', { class: 'hero' }, [
    el('div', { class: 'hero-title' }, `Hola, ${PROFILE.name} 👋`),
    el('div', { class: 'hero-sub' }, `${dayName} · Semana ${week} de 8`),
    el('div', { class: 'streak-pill' }, `🔥 Racha: ${streak} ${streak === 1 ? 'día' : 'días'}`)
  ]));

  const workoutCard = card([
    el('div', { class: 'card-title' }, 'Entrenamiento de hoy'),
    el('div', { class: 'card-line big' }, schedule.type),
    el('div', { class: 'card-line muted' }, schedule.time + (schedule.duration !== '-' ? ' · ' + schedule.duration : '')),
    schedule.note ? el('div', { class: 'card-line warn' }, '⚠ ' + schedule.note) : null,
    schedule.workoutKey
      ? el('button', { class: 'btn primary', onclick: () => { selectedWorkoutKey = schedule.workoutKey; switchTab('entreno'); } }, `Ir a ${WORKOUTS[schedule.workoutKey].label}`)
      : el('div', { class: 'card-line' }, '🚶 Día libre de fuerza. ' + (schedule.day === 'Domingo' ? 'Descanso total.' : 'Caminata o actividad suave si tienes ánimo.'))
  ]);
  frag.appendChild(workoutCard);

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Calorías de hoy'),
    el('div', { class: 'stat-row' }, [
      el('div', { class: 'stat-num' }, `${Math.round(kcal)}`),
      el('div', { class: 'stat-target' }, `/ ${CALORIE_TARGET.min}-${CALORIE_TARGET.max} kcal`)
    ]),
    progressBar(kcal, CALORIE_TARGET.max, kcal > CALORIE_TARGET.max ? 'over' : 'ok'),
    el('div', { class: 'card-title', style: 'margin-top:14px' }, 'Proteína de hoy'),
    el('div', { class: 'stat-row' }, [
      el('div', { class: 'stat-num' }, `${fmt1(prot)} g`),
      el('div', { class: 'stat-target' }, `/ ${PROTEIN_TARGET.min}-${PROTEIN_TARGET.max} g`)
    ]),
    progressBar(prot, PROTEIN_TARGET.max, prot < PROTEIN_TARGET.min ? 'low' : 'ok'),
    el('button', { class: 'btn secondary', onclick: () => switchTab('comidas') }, 'Registrar comida')
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, `Fase actual (semanas ${phase.weeks.join('-')})`),
    el('div', { class: 'card-line' }, phase.desc)
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Hábitos de hoy'),
    habitRow('walk', '🚶 Caminata / cardio', habits),
    habitRow('water', '💧 Buena hidratación', habits),
    el('label', { class: 'habit-row daydone' }, [
      makeCheckbox(habits.dayDone, () => { toggleHabit(today, 'dayDone'); render(); }),
      el('span', {}, '✅ Marcar día como cumplido')
    ])
  ]));

  return frag;
}

function habitRow(key, label, habits) {
  return el('label', { class: 'habit-row' }, [
    makeCheckbox(habits[key], () => { toggleHabit(todayISO(), key); render(); }),
    el('span', {}, label)
  ]);
}

function makeCheckbox(checked, onchange) {
  const cb = el('input', { type: 'checkbox' });
  cb.checked = checked;
  cb.addEventListener('change', onchange);
  return cb;
}

/* ---------- ENTRENO ---------- */

function renderEntreno() {
  const today = todayISO();
  const week = currentWeekNumber();
  const phase = phaseForWeek(week);
  const jsDay = new Date(today + 'T00:00:00').getDay();
  const dayName = DAY_NAMES[jsDay];
  const scheduleToday = WEEKLY_SCHEDULE.find(s => s.day === dayName);
  if (!selectedWorkoutKey) selectedWorkoutKey = scheduleToday.workoutKey || 'A';

  const frag = el('div', {});
  frag.appendChild(el('div', { class: 'tabs-inline' }, ['A', 'B', 'C'].map(k =>
    el('button', {
      class: 'chip' + (k === selectedWorkoutKey ? ' active' : ''),
      onclick: () => { selectedWorkoutKey = k; render(); }
    }, `Fuerza ${k}`)
  )));

  const workout = WORKOUTS[selectedWorkoutKey];
  const wlog = getWorkoutLog(today);
  const isToday = wlog.workoutKey === selectedWorkoutKey;

  frag.appendChild(card([
    el('div', { class: 'card-title' }, workout.label),
    el('div', { class: 'card-line muted' }, workout.focus),
    el('div', { class: 'card-line' }, `Semana ${week}: ${phase.desc}`)
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Antes de partir'),
    el('div', { class: 'card-line' }, WARMUP_COOLDOWN.warmup)
  ]));

  const list = el('div', { class: 'exercise-list' });
  workout.exercises.forEach((ex, idx) => {
    const done = isToday && wlog.completed.includes(idx);
    list.appendChild(el('div', { class: 'exercise-item' + (done ? ' done' : '') }, [
      makeCheckbox(done, () => { toggleExerciseDone(today, selectedWorkoutKey, idx); render(); }),
      el('div', { class: 'exercise-body' }, [
        el('div', { class: 'exercise-name' }, `${ex.name} · ${ex.sets}`),
        el('div', { class: 'exercise-notes' }, ex.notes)
      ])
    ]));
  });
  frag.appendChild(card([el('div', { class: 'card-title' }, 'Ejercicios'), list]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Al terminar'),
    el('div', { class: 'card-line' }, WARMUP_COOLDOWN.cooldown)
  ]));

  return frag;
}

/* ---------- COMIDAS ---------- */

function renderComidas() {
  const today = todayISO();
  const foods = getFoodLog(today);
  const kcal = foods.reduce((s, f) => s + f.kcal, 0);
  const prot = foods.reduce((s, f) => s + f.prot, 0);

  const frag = el('div', {});

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Total de hoy'),
    el('div', { class: 'stat-row' }, [
      el('div', { class: 'stat-num' }, `${Math.round(kcal)} kcal`),
      el('div', { class: 'stat-target' }, `${fmt1(prot)} g proteína`)
    ]),
    progressBar(kcal, CALORIE_TARGET.max, kcal > CALORIE_TARGET.max ? 'over' : 'ok')
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Agregar alimento rápido'),
    quickAddForm(today)
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Agregar alimento manual'),
    manualAddForm(today)
  ]));

  const logList = el('div', { class: 'food-log' });
  if (foods.length === 0) {
    logList.appendChild(el('div', { class: 'muted' }, 'Sin registros hoy todavía.'));
  }
  foods.forEach(f => {
    logList.appendChild(el('div', { class: 'food-item' }, [
      el('div', { class: 'food-item-main' }, [
        el('div', {}, `${f.name}${f.grams ? ' · ' + f.grams + ' g' : ''}`),
        el('div', { class: 'muted small' }, `${Math.round(f.kcal)} kcal · ${fmt1(f.prot)} g prot`)
      ]),
      el('button', { class: 'btn-icon', onclick: () => { removeFoodEntry(today, f.id); render(); } }, '✕')
    ]));
  });
  frag.appendChild(card([el('div', { class: 'card-title' }, 'Registro de hoy'), logList]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Guía de porciones (sin pesar)'),
    portionTable()
  ]));

  return frag;
}

function quickAddForm(dateISO) {
  const select = el('select', { class: 'input' });
  FOOD_DATABASE.forEach(f => select.appendChild(el('option', { value: f.id }, f.name)));
  const gramsInput = el('input', { type: 'number', class: 'input', value: FOOD_DATABASE[0].portion, min: '1' });

  select.addEventListener('change', () => {
    const food = FOOD_DATABASE.find(f => f.id === select.value);
    gramsInput.value = food.portion;
  });

  const addBtn = el('button', {
    class: 'btn primary',
    onclick: () => {
      const food = FOOD_DATABASE.find(f => f.id === select.value);
      const grams = parseFloat(gramsInput.value) || food.portion;
      const kcal = (food.kcal100 * grams) / 100;
      const prot = (food.prot100 * grams) / 100;
      addFoodEntry(dateISO, { name: food.name, grams, kcal, prot });
      render();
    }
  }, 'Agregar');

  return el('div', { class: 'form-row' }, [select, gramsInput, addBtn]);
}

function manualAddForm(dateISO) {
  const nameInput = el('input', { type: 'text', class: 'input', placeholder: 'Nombre del alimento' });
  const kcalInput = el('input', { type: 'number', class: 'input', placeholder: 'kcal', min: '0' });
  const protInput = el('input', { type: 'number', class: 'input', placeholder: 'proteína (g)', min: '0' });
  const addBtn = el('button', {
    class: 'btn secondary',
    onclick: () => {
      if (!nameInput.value.trim()) return;
      addFoodEntry(dateISO, {
        name: nameInput.value.trim(),
        kcal: parseFloat(kcalInput.value) || 0,
        prot: parseFloat(protInput.value) || 0
      });
      render();
    }
  }, 'Agregar');
  return el('div', { class: 'form-row' }, [nameInput, kcalInput, protInput, addBtn]);
}

function portionTable() {
  const table = el('div', { class: 'table' });
  PORTION_GUIDE.forEach(p => {
    table.appendChild(el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, p.tipo),
      el('div', { class: 'table-cell' }, p.ref),
      el('div', { class: 'table-cell muted' }, p.ejemplos)
    ]));
  });
  return table;
}

/* ---------- PROGRESO ---------- */

function renderProgreso() {
  const frag = el('div', {});

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Registrar peso'),
    weightForm()
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Evolución de peso'),
    weightChart(),
    weightList()
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Registrar medidas (cada 2 semanas)'),
    measurementForm()
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Historial de medidas' ),
    measurementList()
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Cómo interpretar los cambios'),
    el('div', { class: 'card-line' }, TRACKING_GUIDE.senalBuena),
    el('div', { class: 'table' }, TRACKING_GUIDE.ajustes.map(a => el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, a.caso),
      el('div', { class: 'table-cell muted' }, a.accion)
    ])))
  ]));

  return frag;
}

function weightForm() {
  const dateInput = el('input', { type: 'date', class: 'input', value: todayISO() });
  const weightInput = el('input', { type: 'number', class: 'input', placeholder: 'Peso (kg)', step: '0.1', min: '20' });
  const btn = el('button', {
    class: 'btn primary',
    onclick: () => {
      const w = parseFloat(weightInput.value);
      if (!w) return;
      addWeightEntry(dateInput.value, w);
      render();
    }
  }, 'Guardar');
  return el('div', { class: 'form-row' }, [dateInput, weightInput, btn]);
}

function weightChart() {
  const log = STORE.weightLog;
  const canvas = el('canvas', { width: '320', height: '140', class: 'chart' });
  if (log.length < 1) {
    return el('div', { class: 'muted' }, 'Aún no hay registros de peso.');
  }
  requestAnimationFrame(() => drawLineChart(canvas, log.map(e => e.weight)));
  return canvas;
}

function drawLineChart(canvas, values) {
  const ctx = canvas.getContext('2d');
  const w = canvas.width, h = canvas.height, pad = 10;
  ctx.clearRect(0, 0, w, h);
  if (values.length < 2) {
    ctx.fillStyle = '#c084b8';
    ctx.beginPath();
    ctx.arc(w / 2, h / 2, 4, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  const min = Math.min(...values), max = Math.max(...values);
  const range = max - min || 1;
  ctx.strokeStyle = '#d16ba5';
  ctx.lineWidth = 2;
  ctx.beginPath();
  values.forEach((v, i) => {
    const x = pad + (i / (values.length - 1)) * (w - pad * 2);
    const y = h - pad - ((v - min) / range) * (h - pad * 2);
    if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
  });
  ctx.stroke();
  ctx.fillStyle = '#d16ba5';
  values.forEach((v, i) => {
    const x = pad + (i / (values.length - 1)) * (w - pad * 2);
    const y = h - pad - ((v - min) / range) * (h - pad * 2);
    ctx.beginPath();
    ctx.arc(x, y, 3, 0, Math.PI * 2);
    ctx.fill();
  });
}

function weightList() {
  const log = [...STORE.weightLog].reverse();
  const list = el('div', { class: 'table' });
  log.forEach((e, i) => {
    const prev = log[i + 1];
    const delta = prev ? fmt1(e.weight - prev.weight) : null;
    list.appendChild(el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell' }, e.date),
      el('div', { class: 'table-cell strong' }, `${e.weight} kg`),
      el('div', { class: 'table-cell muted' }, delta !== null ? (delta > 0 ? `+${delta}` : `${delta}`) + ' kg' : ''),
      el('button', { class: 'btn-icon', onclick: () => { removeWeightEntry(e.date); render(); } }, '✕')
    ]));
  });
  return list;
}

function measurementForm() {
  const dateInput = el('input', { type: 'date', class: 'input', value: todayISO() });
  const cintura = el('input', { type: 'number', class: 'input', placeholder: 'Cintura (cm)', step: '0.1' });
  const cadera = el('input', { type: 'number', class: 'input', placeholder: 'Cadera (cm)', step: '0.1' });
  const muslo = el('input', { type: 'number', class: 'input', placeholder: 'Muslo (cm)', step: '0.1' });
  const gluteo = el('input', { type: 'number', class: 'input', placeholder: 'Glúteo (cm)', step: '0.1' });
  const btn = el('button', {
    class: 'btn primary',
    onclick: () => {
      addMeasurementEntry(dateInput.value, {
        cintura: parseFloat(cintura.value) || null,
        cadera: parseFloat(cadera.value) || null,
        muslo: parseFloat(muslo.value) || null,
        gluteo: parseFloat(gluteo.value) || null
      });
      render();
    }
  }, 'Guardar');
  return el('div', { class: 'form-grid' }, [dateInput, cintura, cadera, muslo, gluteo, btn]);
}

function measurementList() {
  const log = [...STORE.measurementLog].reverse();
  if (log.length === 0) return el('div', { class: 'muted' }, 'Sin medidas registradas.');
  const list = el('div', { class: 'table' });
  log.forEach(e => {
    list.appendChild(el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, e.date),
      el('div', { class: 'table-cell muted' }, `Cintura ${e.cintura ?? '-'} · Cadera ${e.cadera ?? '-'} · Muslo ${e.muslo ?? '-'} · Glúteo ${e.gluteo ?? '-'}`),
      el('button', { class: 'btn-icon', onclick: () => { removeMeasurementEntry(e.date); render(); } }, '✕')
    ]));
  });
  return list;
}

/* ---------- PLAN ---------- */

function renderPlan() {
  const frag = el('div', {});

  frag.appendChild(card([
    el('div', { class: 'card-title' }, `Plan de ${PROFILE.name} · ${PROFILE.age} años · ${PROFILE.heightCm} cm`),
    el('div', { class: 'card-line' }, PROFILE.bmiNote),
    el('div', { class: 'card-line' }, PROFILE.goalNote)
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Cálculo energético y proteína'),
    ...ENERGY_NOTES.map(n => el('div', { class: 'card-line' }, n))
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Semana de entrenamiento' ),
    el('div', { class: 'table' }, WEEKLY_SCHEDULE.map(s => el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, s.day),
      el('div', { class: 'table-cell' }, s.type),
      el('div', { class: 'table-cell muted' }, `${s.time}${s.duration !== '-' ? ' · ' + s.duration : ''}`)
    ])))
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Progresión de 8 semanas'),
    el('div', { class: 'table' }, PROGRESSION_PHASES.map(p => el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, `Semanas ${p.weeks.join('-')}`),
      el('div', { class: 'table-cell muted' }, p.desc)
    ])))
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Cardio y caminata'),
    ...CARDIO_PLAN.map(c => el('div', { class: 'card-line' }, c))
  ]));

  frag.appendChild(equipmentCard('Equipamiento imprescindible', EQUIPMENT.imprescindible));
  frag.appendChild(equipmentCard('Equipamiento recomendado', EQUIPMENT.recomendado));
  frag.appendChild(equipmentCard('Equipamiento opcional', EQUIPMENT.opcional));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Alimentos evitados'),
    el('div', { class: 'card-line' }, FOODS_AVOIDED)
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Menú de 7 días (referencia)'),
    el('div', { class: 'meal-table' }, MEAL_PLAN.map(m => el('div', { class: 'meal-day' }, [
      el('div', { class: 'meal-day-title' }, `Día ${m.day}`),
      el('div', { class: 'meal-row' }, [el('span', { class: 'meal-label' }, 'Desayuno'), el('span', {}, m.desayuno)]),
      el('div', { class: 'meal-row' }, [el('span', { class: 'meal-label' }, 'Almuerzo'), el('span', {}, m.almuerzo)]),
      el('div', { class: 'meal-row' }, [el('span', { class: 'meal-label' }, 'Colación'), el('span', {}, m.colacion)]),
      el('div', { class: 'meal-row' }, [el('span', { class: 'meal-label' }, 'Cena'), el('span', {}, m.cena)])
    ])))
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Seguimiento del progreso'),
    ...TRACKING_GUIDE.frecuencia.map(f => el('div', { class: 'card-line' }, f))
  ]));

  frag.appendChild(card([
    el('div', { class: 'card-title' }, 'Fecha de inicio del plan'),
    startDateForm()
  ]));

  return frag;
}

function equipmentCard(title, items) {
  return card([
    el('div', { class: 'card-title' }, title),
    el('div', { class: 'table' }, items.map(i => el('div', { class: 'table-row' }, [
      el('div', { class: 'table-cell strong' }, i.item),
      el('div', { class: 'table-cell muted' }, `${i.para} · ${i.precio}`),
      el('div', { class: 'table-cell muted small' }, `Alternativa: ${i.alt}`)
    ])))
  ]);
}

function startDateForm() {
  const input = el('input', { type: 'date', class: 'input', value: STORE.settings.startDate });
  const btn = el('button', { class: 'btn secondary', onclick: () => { setStartDate(input.value); render(); } }, 'Guardar');
  return el('div', { class: 'form-row' }, [input, btn]);
}

document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.navbtn').forEach(b => b.addEventListener('click', () => switchTab(b.dataset.tab)));
  render();
});
