const fixtureSocket = new WebSocket('wss://substrate.office.com/m365Copilot/Chathub/fixture?scenario=OfficeWebPaidCopilot');
window.simulateReply = (daily = 32, weekly = 58, absent = false) => {
  fixtureSocket.send(JSON.stringify({ type: 4, target: 'chat', arguments: [{ tone: 'Magic', message: { text: 'Fixture' } }] }) + '\x1e');
  fixtureSocket.receive({ type: 2, item: {
    result: { value: daily === 0 ? 'OutOfCredits' : 'Success' },
    conversationExpiryTime: new Date(Date.now() + 30 * 86_400_000).toISOString(),
    throttling: { numUserMessagesInConversation: 4, maxNumUserMessagesInConversation: 600,
      ...(absent ? {} : { metering: { GPT61SolQueryDaily: { remainingAllowance: daily }, GPT61SolQueryWeekly: { remainingAllowance: weekly } } }) },
    messages: [{ author: 'bot', contentOrigin: daily === 0 ? 'BotConnection' : 'DeepLeo', text: 'Simulated response' }]
  } });
};
document.getElementById('send').addEventListener('click', () => simulateReply());
document.getElementById('cap').addEventListener('click', () => simulateReply(0, 0));
document.getElementById('absent').addEventListener('click', () => simulateReply(32, 58, true));
window.simulateConversation = (used, limit, expiry) => fixtureSocket.receive({ type: 2, item: {
  throttling: { numUserMessagesInConversation: used, maxNumUserMessagesInConversation: limit },
  conversationExpiryTime: expiry
} });
