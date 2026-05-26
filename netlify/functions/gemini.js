exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const { system, userPrompt, maxTokens } = JSON.parse(event.body);
  const API_KEY = process.env.GEMINI_API_KEY;

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        system_instruction: { parts: [{ text: system }] },
        contents: [{ parts: [{ text: userPrompt }] }],
        generationConfig: { maxOutputTokens: maxTokens }
      })
    }
  );

  const data = await response.json();
  if (!response.ok) {
    return { statusCode: 500, body: JSON.stringify({ error: data.error?.message || 'Gemini error' }) };
  }

  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  return {
    statusCode: 200,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ text })
  };
};
