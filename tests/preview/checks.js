const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
document.getElementById('checks').addEventListener('click', async () => {
  const result = document.getElementById('results'); const passed = [];
  const check = (condition, name) => { if (!condition) throw new Error(name); passed.push(name); };
  result.textContent = 'Running…';
  try {
    const copilot = document.getElementById('copilot').contentWindow;
    const popup = document.getElementById('popup').contentWindow;
    const host = copilot.document.querySelector('m365-copilot-extender');
    const pageRoot = host.shadowRoot;
    const popupRoot = popup.document.getElementById('panel').shadowRoot;
    const choose = (root, value) => { const select = root.querySelector('select'); select.value = value; select.dispatchEvent(new Event('change', { bubbles: true })); };
    await fixtureStore.set({ selectedTone: 'Gpt_61_Sol_Reasoning', inPageEnabled: true }); await pause(30);
    check(pageRoot.querySelectorAll('option').length === 6, 'Six model options');
    choose(popupRoot, 'Gpt_6_Sol_Reasoning'); await pause(30);
    check(pageRoot.querySelector('select').value === 'Gpt_6_Sol_Reasoning', 'Popup selection reaches page');
    check(pageRoot.querySelector('.usage').hidden, 'No invented GPT-6 Sol quota');
    copilot.simulateReply(); await pause(30);
    check(copilot.document.getElementById('wire').textContent.includes('Gpt_6_Sol_Reasoning'), 'Outgoing request is overridden');
    choose(pageRoot, 'Gpt_61_Sol_Reasoning'); await pause(30);
    check(popupRoot.querySelector('select').value === 'Gpt_61_Sol_Reasoning', 'Page selection reaches popup');
    copilot.simulateReply(0, 0); await pause(30);
    check(pageRoot.querySelector('.daily').textContent === '0', 'Zero quota stays visible');
    check(pageRoot.querySelector('.status-text').textContent.includes('usage limit'), 'Quota refusal is surfaced');
    copilot.simulateReply(30, 50, true); await pause(30);
    check(pageRoot.querySelector('.counts').hidden && !pageRoot.querySelector('.pending').hidden, 'Missing quota stays unknown');
    const toggle = popupRoot.querySelector('input'); toggle.checked = false; toggle.dispatchEvent(new Event('change')); await pause(30);
    check(host.style.display === 'none', 'Popup hides the page panel');
    check(fixtureStore.data.inPageEnabled === false, 'Panel visibility is saved');
    toggle.checked = true; toggle.dispatchEvent(new Event('change')); await pause(30);
    check(host.style.display === 'block', 'Popup restores the page panel');
    const collapse = pageRoot.querySelector('.collapse'); collapse.click();
    check(pageRoot.querySelector('.card').classList.contains('compact'), 'Page panel collapses'); collapse.click();
    copilot.simulateReply(32, 58); await pause(1600);
    check(popupRoot.querySelector('.daily').textContent === '32', 'Popup reads active-tab quota');
    result.textContent = passed.length + ' UI checks passed · simulated transport'; result.dataset.status = 'passed';
  } catch (error) { result.textContent = 'FAILED: ' + error.message; result.dataset.status = 'failed'; }
});
