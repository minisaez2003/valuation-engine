// Vercel/Netlify serverless function: /api/chat
// Add ANTHROPIC_API_KEY to your environment variables in Vercel dashboard.
// This keeps your API key server-side so it's never exposed in the browser.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  const apiKey = process.env.ANTHROPIC_API_KEY
  if (!apiKey) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured on server. Add it in your deployment environment variables.' })
  }

  try {
    const { system, messages } = req.body
    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: 'messages array required' })
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1024,
        system: system || 'You are a helpful M&A analyst assistant.',
        messages,
      }),
    })

    if (!response.ok) {
      const errText = await response.text()
      return res.status(response.status).json({ error: `Anthropic API error: ${errText}` })
    }

    const data = await response.json()
    const reply = data.content?.[0]?.text || '(no response)'
    return res.status(200).json({ reply })
  } catch (e) {
    return res.status(500).json({ error: e.message })
  }
}
