function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const AI_PROMPT = `Eres un nutricionista estimando calorías a partir de una foto de comida casera chilena.
Identifica cada alimento visible en el plato y estima su porción en gramos usando referencias visuales (una palma de mano ~150-180 g de proteína, un puño ~150 g de carbohidrato cocido, un pulgar ~15 g de grasa/salsa).
Responde ÚNICAMENTE con un JSON válido (sin texto adicional, sin markdown) con esta forma exacta:
{"items":[{"name":"string","grams":number,"kcal":number,"prot":number}],"notes":"string breve en español, opcional"}
Si no puedes identificar comida en la imagen, responde {"items":[],"notes":"No se detectó comida en la foto."}`;

async function analyzeFoodPhoto(file) {
  const apiKey = STORE.settings.apiKey;
  if (!apiKey) {
    throw new Error('Falta configurar tu API key de Anthropic en Ajustes.');
  }
  const base64 = await fileToBase64(file);
  const mediaType = file.type || 'image/jpeg';

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: STORE.settings.aiModel || 'claude-sonnet-5',
      max_tokens: 1024,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
          { type: 'text', text: AI_PROMPT }
        ]
      }]
    })
  });

  if (!response.ok) {
    const errBody = await response.text().catch(() => '');
    if (response.status === 401) throw new Error('API key inválida. Revísala en Ajustes.');
    if (response.status === 429) throw new Error('Se alcanzó el límite de uso de la API. Intenta de nuevo más tarde.');
    throw new Error('Error consultando la IA (' + response.status + '). ' + errBody.slice(0, 200));
  }

  const data = await response.json();
  const text = (data.content || []).map(b => b.text || '').join('').trim();
  let parsed;
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
  } catch (e) {
    throw new Error('No se pudo interpretar la respuesta de la IA.');
  }
  if (!parsed.items) parsed.items = [];
  return parsed;
}
