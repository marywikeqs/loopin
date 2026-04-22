exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  let channel, message;
  try {
    ({ channel, message } = JSON.parse(event.body));
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request' }) };
  }

  if (!message || !message.trim()) {
    return { statusCode: 400, body: JSON.stringify({ error: 'Message is required' }) };
  }

  // Map channel keys to environment variables
  // To add a new channel: add an entry here and set the env var in Netlify
  const webhooks = {
    test:            process.env.SLACK_WEBHOOK_TEST,
    sales_leadership: process.env.SLACK_WEBHOOK_SALES_LEADERSHIP,
  };

  const webhook = webhooks[channel];
  if (!webhook) {
    return { statusCode: 400, body: JSON.stringify({
      error: 'Unknown channel',
      received: channel,
      env_test_set: !!process.env.SLACK_WEBHOOK_TEST,
    })};
  }

  try {
    const res = await fetch(webhook, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: message }),
    });

    if (res.ok) {
      return { statusCode: 200, body: JSON.stringify({ ok: true }) };
    } else {
      return { statusCode: 500, body: JSON.stringify({ error: 'Slack returned an error' }) };
    }
  } catch (e) {
    return { statusCode: 500, body: JSON.stringify({ error: 'Failed to reach Slack' }) };
  }
};
