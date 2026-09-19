const PROFILE = {
  name: 'Cande',
  age: 22,
  heightCm: 155,
  startWeightKg: 50,
  bmiNote: 'Con 50 kg y 1,55 m el IMC es 20,8 (rango saludable). Bajar a 45 kg dejaría el IMC en 18,7, en el límite inferior. No es peligroso, pero da poco margen si a la vez se busca ganar músculo.',
  goalNote: 'El foco es bajar grasa y ganar tono manteniendo la mayor masa muscular posible, dejando que el peso final se acomode según cómo responda el cuerpo. Si el peso se estabiliza en 47-48 kg con más músculo y menos grasa, ese resultado es mejor que forzar el número 45.'
};

const CALORIE_TARGET = { min: 1400, max: 1500, hardFloor: 1300 };
const PROTEIN_TARGET = { min: 90, max: 100 };

const ENERGY_NOTES = [
  'Gasto energético basal estimado (Mifflin-St Jeor, 22 años / 155 cm / 50 kg): ≈ 1.200 kcal en reposo.',
  'Sumando la actividad diaria y las 3 sesiones de fuerza semanales, el gasto total estimado es ≈ 1.600 a 1.650 kcal/día.',
  'Punto de partida sugerido: 1.400 a 1.500 kcal/día (déficit moderado). Nunca bajar de 1.300 kcal.',
  'Proteína objetivo: 90 a 100 g/día. Es la prioridad de la alimentación, más que contar calorías al detalle.',
  'Ajuste: si en 3-4 semanas el peso y las medidas no cambian, bajar a 1.300-1.400 kcal. Si hay cansancio, mal rendimiento u hambre constante, subir a 1.500-1.600 kcal en vez de seguir bajando.'
];

const WEEKLY_SCHEDULE = [
  { day: 'Lunes', type: 'Fuerza A', workoutKey: 'A', time: 'Desde las 18:00', duration: '45 a 50 min' },
  { day: 'Martes', type: 'Descanso activo', workoutKey: null, time: 'Caminata a mediodía o tarde', duration: '20 a 30 min' },
  { day: 'Miércoles', type: 'Fuerza B', workoutKey: 'B', time: 'Después de las 12:00', duration: '45 a 50 min' },
  { day: 'Jueves', type: 'Descanso o caminata opcional', workoutKey: null, time: 'Cuando prefieras', duration: '0 a 30 min' },
  { day: 'Viernes', type: 'Fuerza C', workoutKey: 'C', time: 'Antes de las 17:00 o después de las 18:00', duration: '45 a 50 min', note: 'Dejar libre el bloque 17:00-18:00 por la hora con el psicólogo.' },
  { day: 'Sábado', type: 'Actividad libre opcional', workoutKey: null, time: 'Cuando prefieras', duration: '20 a 40 min' },
  { day: 'Domingo', type: 'Descanso completo', workoutKey: null, time: '-', duration: '-' }
];

const PROGRESSION_PHASES = [
  { weeks: [1, 2], desc: 'Aprender la técnica de cada ejercicio. 2 series por ejercicio, peso corporal o muy liviano. Prioridad: moverse bien, no cansarse mucho.' },
  { weeks: [3, 4, 5], desc: '3 series por ejercicio. Sumar peso liviano (mancuernas o botellas). Descansos de 45 a 60 segundos.' },
  { weeks: [6, 7, 8], desc: 'Subir el peso o cambiar a variantes más exigentes (a una pierna, tempo más lento, más repeticiones). Descansos de 60 a 90 segundos en ejercicios de pierna y glúteo.' }
];

function phaseForWeek(week) {
  return PROGRESSION_PHASES.find(p => p.weeks.includes(week)) || PROGRESSION_PHASES[PROGRESSION_PHASES.length - 1];
}

const WARMUP_COOLDOWN = {
  warmup: '5 minutos de movilidad y activación: rotaciones de cadera, sentadillas sin peso, marcha en el lugar.',
  cooldown: '5 minutos de estiramiento de piernas, glúteo y espalda.'
};

const WORKOUTS = {
  A: {
    label: 'Fuerza A · Lunes',
    focus: 'Piernas y glúteo',
    exercises: [
      { name: 'Sentadilla goblet', sets: '3 x 12-15', notes: 'Pies al ancho de hombros, bajar como si te sentaras en una silla, pecho arriba. Error común: rodillas hacia adentro. Progresión: sumar peso o bajar más lento.' },
      { name: 'Puente de glúteo', sets: '3 x 15', notes: 'Acostada, pies cerca de los glúteos, subir la cadera apretando el glúteo arriba. Progresión: hacerlo a una pierna.' },
      { name: 'Zancada estática (split squat)', sets: '3 x 10 por lado', notes: 'Un pie adelante y uno atrás, bajar la rodilla trasera casi al piso. Error común: paso muy corto. Progresión: sumar peso en las manos.' },
      { name: 'Plancha abdominal', sets: '3 x 20-30 seg', notes: 'Apoyo en antebrazos, cuerpo recto de cabeza a talones, sin dejar caer la cadera. Progresión: aumentar el tiempo.' },
      { name: 'Remo con banda elástica', sets: '3 x 12', notes: 'Banda anclada adelante, tirar los codos hacia atrás juntando los omóplatos. Progresión: banda más firme.' },
      { name: 'Elevación de talones', sets: '3 x 15', notes: 'De pie, subir los talones lo más arriba posible y bajar controlado. Progresión: hacerlo a una pierna.' }
    ]
  },
  B: {
    label: 'Fuerza B · Miércoles',
    focus: 'Glúteo medio y core',
    exercises: [
      { name: 'Sentadilla sumo', sets: '3 x 12', notes: 'Pies más separados que el ancho de hombros, puntas hacia afuera, bajar recto. Trabaja glúteo e interior de pierna.' },
      { name: 'Patada de glúteo (donkey kick)', sets: '3 x 15 por lado', notes: 'En cuatro apoyos, llevar una pierna hacia atrás y arriba sin arquear la espalda baja.' },
      { name: 'Abducción de cadera con banda', sets: '3 x 15 por lado', notes: 'Banda sobre las rodillas, de pie o tumbada de lado, separar la pierna contra la resistencia de la banda.' },
      { name: 'Peso muerto rumano', sets: '3 x 10', notes: 'Con mancuernas o botellas, piernas casi rectas, bajar el peso pegado a las piernas sintiendo el estiramiento en la parte de atrás del muslo.' },
      { name: 'Dead bug', sets: '3 x 10 por lado', notes: 'Acostada boca arriba, bajar brazo y pierna contrarios sin despegar la espalda baja del piso.' },
      { name: 'Superman', sets: '3 x 12', notes: 'Boca abajo, levantar brazos y piernas al mismo tiempo apretando la espalda baja y glúteo.' }
    ]
  },
  C: {
    label: 'Fuerza C · Viernes',
    focus: 'Cuerpo completo',
    exercises: [
      { name: 'Sentadilla con salto suave', sets: '3 x 12', notes: 'En semanas 1 y 2 hacer sentadilla normal sin salto, se agrega el salto desde la semana 3 si el cuerpo lo permite bien.' },
      { name: 'Zancada caminando', sets: '3 x 10 por lado', notes: 'Dar pasos largos alternando piernas, bajando la rodilla trasera cerca del piso en cada paso.' },
      { name: 'Puente de glúteo a una pierna', sets: '3 x 10 por lado', notes: 'Igual que el puente normal pero con una pierna estirada al aire.' },
      { name: 'Press de hombro', sets: '3 x 12', notes: 'Con mancuernas o botellas, empujar desde los hombros hacia arriba sin arquear la espalda.' },
      { name: 'Plancha lateral', sets: '3 x 20 seg por lado', notes: 'Apoyo en un antebrazo, cuerpo en línea recta, cadera arriba sin caer.' },
      { name: 'Curl de bíceps', sets: '3 x 12', notes: 'Con banda o mancuernas, subir el peso doblando el codo sin balancear el cuerpo.' }
    ]
  }
};

const EQUIPMENT = {
  imprescindible: [
    { item: 'Colchoneta de ejercicio', para: 'Comodidad para ejercicios en el piso', precio: '$10.000 a $18.000', alt: 'Una toalla gruesa o manta doblada' },
    { item: 'Set de bandas elásticas de resistencia', para: 'Trabajar glúteo y espalda con resistencia progresiva', precio: '$10.000 a $15.000', alt: 'Ninguna buena, conviene comprarlas' },
    { item: 'Par de mancuernas livianas (2 a 4 kg)', para: 'Sumar peso a sentadillas, zancadas, remo y press', precio: '$15.000 a $25.000 el par', alt: 'Botellas de bebida o agua rellenas' }
  ],
  recomendado: [
    { item: 'Miniband de glúteo (banda circular)', para: 'Trabajo específico de glúteo medio', precio: '$5.000 a $8.000', alt: 'Banda elástica larga anudada en círculo' },
    { item: 'Segundo par de mancuernas (6 a 8 kg)', para: 'Progresar cuando el peso liviano ya sea fácil', precio: '$20.000 a $35.000 el par', alt: 'Bolsas o mochila con peso adentro' }
  ],
  opcional: [
    { item: 'Step o cajón bajo', para: 'Subidas y ejercicios de pierna con más rango', precio: '$15.000 a $30.000', alt: 'Un escalón de la casa' },
    { item: 'Cronómetro o app de intervalos', para: 'Controlar tiempos de descanso y plancha', precio: 'Gratis en el celular', alt: '-' }
  ]
};

const CARDIO_PLAN = [
  'Martes y jueves: caminata de 20 a 30 minutos a paso moderado (se puede dividir en dos tramos de 10 a 15 min).',
  'Sábado o domingo: una actividad libre y entretenida de 20 a 40 minutos (caminar, bailar, bicicleta).',
  'No es necesario cardio los días de fuerza, el entrenamiento ya es suficiente estímulo para ese día.',
  'Progresión: si en 2-3 semanas el cuerpo se siente cómodo, sumar más caminata diaria (subir escaleras, caminar tramos que antes se hacían en auto).'
];

const PORTION_GUIDE = [
  { tipo: 'Proteína (por comida)', ref: '1 palma de mano de grosor y tamaño', ejemplos: '150 a 180 g de pollo o carne, 2 a 3 huevos' },
  { tipo: 'Carbohidrato cocido (por comida)', ref: '1 puño cerrado a 1 taza', ejemplos: 'Arroz, pasta, papas, legumbres' },
  { tipo: 'Grasas extra', ref: '1 pulgar o 1 cucharada', ejemplos: 'Palta, mantequilla de maní, aceite' },
  { tipo: 'Fruta', ref: '1 puño o 1 unidad mediana', ejemplos: 'Plátano, manzana, taza de fruta picada' }
];

const FOODS_AVOIDED = 'pescados, mariscos, brócoli, espinaca, acelga, champiñones, zapallo italiano, aceitunas, betarraga, repollo, interiores, huevo duro y quesos que no sean derretidos o parmesano';

const MEAL_PLAN = [
  { day: 1, desayuno: '1/2 taza de avena cruda con 1 taza de leche, 1 plátano y 1 cucharada de mantequilla de maní', almuerzo: '1 palma de pollo a la plancha, 1 taza de arroz, 1/2 palta y tomate', colacion: '1 yogur natural con fruta', cena: '2 a 3 tacos de carne molida (150 g) con lechuga y tomate' },
  { day: 2, desayuno: '2 huevos revueltos con 1 a 2 rebanadas de pan y 1 fruta', almuerzo: '1 palma de carne al jugo, 2 papas medianas cocidas y ensalada de tomate', colacion: '1 fruta con un puñado de almendras', cena: '1 taza de pasta con salsa de tomate y pollo desmenuzado (del día anterior)' },
  { day: 3, desayuno: '1 yogur con 3 cucharadas de avena y fruta picada', almuerzo: '1 taza de garbanzos guisados con 1/2 taza de arroz y trozos de pollo', colacion: '1 vaso de leche con fruta', cena: 'Sandwich de pollo o carne con palta, más ensalada al lado' },
  { day: 4, desayuno: 'Pan con palta y 1 a 2 huevos fritos o pochados', almuerzo: '1 palma de pollo al horno (junto con las papas) con papas y ensalada', colacion: '1 yogur natural', cena: 'Arroz salteado con huevo y las verduras que te gusten' },
  { day: 5, desayuno: '1/2 taza de avena con leche y fruta', almuerzo: '1 taza de butter chicken con 1 taza de arroz', colacion: 'Tostada con queso derretido', cena: 'Tacos de pollo, usando el pollo que sobró del almuerzo' },
  { day: 6, desayuno: '2 huevos a la copa con pan', almuerzo: '1 palma de carne asada con papas y ensalada (dejar el doble para el otro día)', colacion: '1 yogur con fruta', cena: '1 taza de pasta con el pollo o carne que quedó lista del fin de semana' },
  { day: 7, desayuno: '2 a 3 panqueques de avena con fruta y yogur', almuerzo: 'Lasaña de carne, 1 trozo mediano, con ensalada', colacion: '1 fruta', cena: 'Algo liviano y rápido: huevos revueltos con pan o un sandwich' }
];

const TRACKING_GUIDE = {
  frecuencia: [
    'Peso: 1 vez por semana, mismo día y en ayunas (no todos los días).',
    'Medidas: cintura, cadera, muslo y contorno de glúteo cada 2 semanas con huincha.',
    'Fotos de progreso cada 2 a 4 semanas, misma luz y misma pose.',
    'Fuerza: anotar peso y repeticiones de cada ejercicio para ver si van subiendo semana a semana.',
    'Energía y sueño: registro simple de cómo te sientes día a día.'
  ],
  senalBuena: 'El peso baja lento (no más de 0,3 a 0,5 kg por semana), las medidas de cintura bajan, y la fuerza en los ejercicios se mantiene o sube.',
  ajustes: [
    { caso: 'Bajas muy rápido (más de 0,5 kg/semana varias semanas seguidas)', accion: 'Subir la comida, el déficit está siendo muy fuerte.' },
    { caso: 'No bajas nada en 3-4 semanas y las medidas tampoco cambian', accion: 'Bajar un poco las calorías o revisar si las porciones han ido creciendo.' },
    { caso: 'Tienes mucha hambre todo el día', accion: 'Subir la proteína y verduras/legumbres que te gusten, no solo aguantar.' },
    { caso: 'Estás sin energía o rindes mal entrenando', accion: 'Subir un poco las calorías antes de seguir bajando.' },
    { caso: 'El peso sube pero las medidas bajan', accion: 'Buena señal, probablemente ganando músculo, seguir igual.' },
    { caso: 'Un ejercicio causa dolor articular o agudo', accion: 'Parar ese ejercicio, revisar técnica y probar una versión más simple.' }
  ]
};

const FOOD_DATABASE = [
  { id: 'pollo', name: 'Pollo (pechuga cocida)', kcal100: 165, prot100: 31, portion: 160 },
  { id: 'carne', name: 'Carne de vacuno (asada/molida)', kcal100: 250, prot100: 26, portion: 150 },
  { id: 'huevo', name: 'Huevo (unidad, ~50 g)', kcal100: 155, prot100: 13, portion: 50 },
  { id: 'arroz', name: 'Arroz cocido', kcal100: 130, prot100: 2.7, portion: 150 },
  { id: 'pasta', name: 'Pasta cocida', kcal100: 158, prot100: 5.8, portion: 150 },
  { id: 'papa', name: 'Papa cocida', kcal100: 87, prot100: 2, portion: 200 },
  { id: 'avena', name: 'Avena cruda', kcal100: 389, prot100: 16.9, portion: 40 },
  { id: 'pan', name: 'Pan (rebanada)', kcal100: 265, prot100: 9, portion: 30 },
  { id: 'palta', name: 'Palta', kcal100: 160, prot100: 2, portion: 70 },
  { id: 'yogur', name: 'Yogur natural', kcal100: 61, prot100: 3.5, portion: 170 },
  { id: 'leche', name: 'Leche', kcal100: 42, prot100: 3.4, portion: 250 },
  { id: 'platano', name: 'Plátano', kcal100: 89, prot100: 1.1, portion: 120 },
  { id: 'manzana', name: 'Manzana', kcal100: 52, prot100: 0.3, portion: 150 },
  { id: 'almendras', name: 'Almendras (puñado)', kcal100: 579, prot100: 21, portion: 20 },
  { id: 'mantmani', name: 'Mantequilla de maní', kcal100: 588, prot100: 25, portion: 15 },
  { id: 'garbanzos', name: 'Garbanzos cocidos', kcal100: 164, prot100: 8.9, portion: 150 },
  { id: 'queso', name: 'Queso (derretido/parmesano)', kcal100: 380, prot100: 28, portion: 30 },
  { id: 'aceite', name: 'Aceite / mantequilla', kcal100: 884, prot100: 0, portion: 10 },
  { id: 'lechetostada', name: 'Tostada con queso derretido', kcal100: 300, prot100: 12, portion: 60 }
];

const DAY_NAMES = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
