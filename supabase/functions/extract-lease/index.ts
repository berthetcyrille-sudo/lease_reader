Deno.serve(async (req) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  }

  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const { fileUrl, mediaType, prompt } = await req.json()

    if (!fileUrl || !mediaType || !prompt) {
      return new Response(
        JSON.stringify({ error: 'Paramètres manquants (fileUrl, mediaType, prompt)' }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    const ANTHROPIC_API_KEY = Deno.env.get('ANTHROPIC_API_KEY') ?? ''

    // Passer l'URL directement à Claude — pas de téléchargement ni de base64
    const anthropicRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-opus-4-5',
        max_tokens: 4096,
        messages: [
          {
            role: 'user',
            content: [
              {
                type: 'document',
                source: { type: 'url', url: fileUrl },
              },
              { type: 'text', text: prompt },
            ],
          },
        ],
      }),
    })

    if (!anthropicRes.ok) {
      const errText = await anthropicRes.text()
      return new Response(
        JSON.stringify({ error: 'Anthropic API error: ' + errText }),
        { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    const anthropicData = await anthropicRes.json()
    const raw = anthropicData.content
      .map((b) => b.text ?? '')
      .join('')
      .trim()
      .replace(/^```json\s*/, '')
      .replace(/\s*```$/, '')
      .trim()

    let result
    try {
      result = JSON.parse(raw)
    } catch (_) {
      return new Response(
        JSON.stringify({ error: 'Réponse Claude non parseable en JSON', raw }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    return new Response(
      JSON.stringify({ result }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  } catch (e) {
    return new Response(
      JSON.stringify({ error: String(e) }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})
