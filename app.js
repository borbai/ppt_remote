/**
 * PPT Remote Web Client
 * Conecta-se ao canal Supabase Realtime para controlar a apresentação e receber dados em tempo real.
 */

// Estado da Aplicação
let supabaseClient = null;
let realtimeChannel = null;
let currentConfig = {
  supabaseUrl: '',
  supabaseKey: '',
  roomId: ''
};

let presentationState = {
  isConnected: false,
  isPresenting: false,
  name: '',
  currentSlide: 0,
  totalSlides: 0,
  title: '',
  notes: '',
  screenState: 'normal'
};

// Cronômetro
let timerSeconds = 0;
let timerInterval = null;
let timerRunning = false;

// Tamanho da fonte das notas
let notesFontSize = 0.95;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
  loadConfigFromStorageOrUrl();
  setupKeyboardShortcuts();
  
  if (currentConfig.supabaseUrl && currentConfig.supabaseKey && currentConfig.roomId) {
    connectToSupabase();
  } else {
    toggleConfigModal(true);
  }
});

// Leitura de parâmetros de URL e LocalStorage
function loadConfigFromStorageOrUrl() {
  // Tenta ler da URL (Hash ou Query Params)
  const urlParams = new URLSearchParams(window.location.search);
  const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ''));

  const urlParam = (key) => urlParams.get(key) || hashParams.get(key);

  const urlRoom = urlParam('room');
  const urlSupabaseUrl = urlParam('url');
  const urlSupabaseKey = urlParam('key');

  const savedUrl = localStorage.getItem('ppt_remote_url') || '';
  const savedKey = localStorage.getItem('ppt_remote_key') || '';
  const savedRoom = localStorage.getItem('ppt_remote_room') || '';

  currentConfig.supabaseUrl = urlSupabaseUrl || savedUrl;
  currentConfig.supabaseKey = urlSupabaseKey || savedKey;
  currentConfig.roomId = urlRoom || savedRoom;

  // Preenche inputs do modal
  document.getElementById('supabaseUrlInput').value = currentConfig.supabaseUrl;
  document.getElementById('supabaseKeyInput').value = currentConfig.supabaseKey;
  document.getElementById('roomIdInput').value = currentConfig.roomId;
}

// Conectar ao Supabase Realtime
function connectToSupabase() {
  const { supabaseUrl, supabaseKey, roomId } = currentConfig;

  if (!supabaseUrl || !supabaseKey || !roomId) {
    showToast('Preencha os dados de conexão!');
    toggleConfigModal(true);
    return;
  }

  updateConnectionStatus('connecting', 'Conectando...');

  try {
    // Se houver canal ou cliente anterior, limpa
    if (realtimeChannel) {
      realtimeChannel.unsubscribe();
    }

    supabaseClient = window.supabase.createClient(supabaseUrl, supabaseKey);

    const channelName = `room_${roomId}`;
    realtimeChannel = supabaseClient.channel(channelName, {
      config: {
        broadcast: { ack: false, self: false }
      }
    });

    // Escuta atualizações de estado do PowerPoint enviadas pelo PC
    realtimeChannel.on('broadcast', { event: 'state_update' }, (message) => {
      handleStateUpdate(message.payload);
    });

    // Subscreve ao canal
    realtimeChannel.subscribe((status) => {
      if (status === 'SUBSCRIBED') {
        updateConnectionStatus('connected', `Sala: ${roomId}`);
        showToast(`Conectado à sala ${roomId}!`);
        
        // Solicita o estado atual da apresentação ao host Windows
        sendCommand('request_state');
      } else if (status === 'CLOSED' || status === 'CHANNEL_ERROR') {
        updateConnectionStatus('disconnected', 'Desconectado');
      }
    });

    // Salva no LocalStorage para os próximos acessos
    localStorage.setItem('ppt_remote_url', supabaseUrl);
    localStorage.setItem('ppt_remote_key', supabaseKey);
    localStorage.setItem('ppt_remote_room', roomId);

  } catch (err) {
    console.error('Erro ao conectar ao Supabase:', err);
    updateConnectionStatus('disconnected', 'Erro');
    showError(err.message || 'Falha ao conectar.');
  }
}

// Atualiza o badge visual de conexão
function updateConnectionStatus(type, text) {
  const badge = document.getElementById('connectionBadge');
  const textEl = document.getElementById('connectionText');
  badge.className = `connection-badge ${type}`;
  textEl.innerText = text;
}

// Enviar Comando Broadcast
function sendCommand(action, payload = {}) {
  triggerHaptic();

  if (!realtimeChannel) {
    showToast('Não conectado à sala!');
    return;
  }

  realtimeChannel.send({
    type: 'broadcast',
    event: 'command',
    payload: {
      action: action,
      ...payload
    }
  });
}

// Resposta com estado vinda do PC Windows
function handleStateUpdate(state) {
  if (!state) return;

  presentationState = { ...presentationState, ...state };

  // Atualiza Seletor ou Nome da Apresentação
  const presNameEl = document.getElementById('presName');
  const presSelectEl = document.getElementById('presSelect');
  const presentations = state.presentations || [];

  if (presentations.length > 1) {
    presNameEl.classList.add('hidden');
    presSelectEl.classList.remove('hidden');

    // Reconstrói as opções caso tenham mudado
    const currentOptions = Array.from(presSelectEl.options).map(o => o.value).join('|');
    const newOptions = presentations.map(p => p.name).join('|');

    if (currentOptions !== newOptions) {
      presSelectEl.innerHTML = '';
      presentations.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.name;
        opt.textContent = `${p.name} ${p.is_presenting ? '🟢 (Tela cheia)' : ''}`;
        presSelectEl.appendChild(opt);
      });
    }

    const selectedName = state.selected_presentation || state.presentation_name;
    if (selectedName && presSelectEl.value !== selectedName) {
      presSelectEl.value = selectedName;
    }
  } else {
    presSelectEl.classList.add('hidden');
    presNameEl.classList.remove('hidden');

    if (state.presentation_name) {
      presNameEl.innerText = state.presentation_name;
      presNameEl.title = state.presentation_name;
    } else if (state.is_connected) {
      presNameEl.innerText = 'PowerPoint aberto (sem apresentação ativa)';
    } else {
      presNameEl.innerText = 'Aguardando PowerPoint...';
    }
  }

  // Contador de Slides
  const currentSlideEl = document.getElementById('currentSlideNum');
  const totalSlidesEl = document.getElementById('totalSlidesNum');
  const progressBar = document.getElementById('slideProgressBar');

  currentSlideEl.innerText = state.current_slide || 0;
  totalSlidesEl.innerText = state.total_slides || 0;

  if (state.total_slides > 0) {
    const pct = Math.min(100, Math.max(0, (state.current_slide / state.total_slides) * 100));
    progressBar.style.width = `${pct}%`;
  } else {
    progressBar.style.width = '0%';
  }

  // Título do Slide
  const titleEl = document.getElementById('slideTitle');
  if (state.slide_title) {
    titleEl.innerText = state.slide_title;
  } else if (state.is_presenting) {
    titleEl.innerText = `Slide ${state.current_slide}`;
  } else {
    titleEl.innerText = 'Nenhuma apresentação em tela cheia';
  }

  // Miniaturas e Detecção de Animação nos Botões
  updateButtonsWithThumbsAndAnim(state);

  // Anotações do Orador
  const notesEl = document.getElementById('speakerNotesText');
  if (state.speaker_notes && state.speaker_notes.trim()) {
    notesEl.innerText = state.speaker_notes;
  } else {
    notesEl.innerText = 'Sem anotações para este slide.';
  }

  // Estado dos botões de ação secundária
  const btnBlack = document.getElementById('btnBlack');
  const btnWhite = document.getElementById('btnWhite');
  const btnStartStop = document.getElementById('btnStartStop');
  const startStopIcon = document.getElementById('startStopIcon');
  const startStopText = document.getElementById('startStopText');

  if (state.screen_state === 'black') {
    btnBlack.style.borderColor = 'var(--status-warning)';
  } else {
    btnBlack.style.borderColor = 'var(--border-color)';
  }

  if (state.screen_state === 'white') {
    btnWhite.style.borderColor = 'var(--status-warning)';
  } else {
    btnWhite.style.borderColor = 'var(--border-color)';
  }

  if (state.is_presenting) {
    startStopIcon.innerText = '⏹️';
    startStopText.innerText = 'Encerrar';
  } else {
    startStopIcon.innerText = '▶️';
    startStopText.innerText = 'Apresentar';
  }
}

// Atualiza visual dos botões de avançar/voltar com miniaturas e animações
function updateButtonsWithThumbsAndAnim(state) {
  const btnPrev = document.getElementById('btnPrev');
  const prevThumbImg = document.getElementById('prevThumbImg');
  const prevBadge = document.getElementById('prevBadge');
  const prevTitle = document.getElementById('prevBtnTitle');
  const prevSub = document.getElementById('prevBtnSub');

  const btnNext = document.getElementById('btnNext');
  const nextThumbImg = document.getElementById('nextThumbImg');
  const nextBadge = document.getElementById('nextBadge');
  const nextTitle = document.getElementById('nextBtnTitle');
  const nextSub = document.getElementById('nextBtnSub');

  const currentSlide = state.current_slide || 0;
  const totalSlides = state.total_slides || 0;

  // Botão Anterior
  if (state.prev_is_animation) {
    prevBadge.classList.remove('hidden');
    prevBadge.innerText = '↩️ Desfazer Animação';
    prevTitle.innerText = 'Voltar';
    prevSub.innerText = 'Passo Anterior';
  } else {
    if (currentSlide > 1) {
      prevBadge.classList.remove('hidden');
      prevBadge.innerText = `Slide ${currentSlide - 1}`;
      prevTitle.innerText = 'Anterior';
      prevSub.innerText = `Voltar para Slide ${currentSlide - 1}`;
    } else {
      prevBadge.classList.add('hidden');
      prevTitle.innerText = 'Anterior';
      prevSub.innerText = 'Início';
    }
  }

  if (state.prev_thumb) {
    prevThumbImg.src = state.prev_thumb;
    prevThumbImg.classList.remove('hidden');
  } else {
    prevThumbImg.classList.add('hidden');
  }

  // Botão Próximo
  if (state.next_is_animation) {
    // DESTACAR QUE O PRÓXIMO CLIQUE AVANÇA UMA ANIMAÇÃO
    btnNext.classList.add('is-animation');
    nextBadge.classList.remove('hidden');

    const step = (state.animation_current_step || 0) + 1;
    const totalSteps = state.animation_total_steps || 1;
    nextBadge.innerText = `⚡ Próxima: Animação (${step}/${totalSteps})`;

    nextTitle.innerText = '⚡ Animação';
    nextSub.innerText = `Avançar Passo ${step} de ${totalSteps}`;

    // Mostra o slide atual onde a animação será exibida
    const animThumb = state.current_thumb || state.next_thumb;
    if (animThumb) {
      nextThumbImg.src = animThumb;
      nextThumbImg.classList.remove('hidden');
    } else {
      nextThumbImg.classList.add('hidden');
    }
  } else {
    btnNext.classList.remove('is-animation');

    if (currentSlide < totalSlides) {
      nextBadge.classList.remove('hidden');
      nextBadge.innerText = `Slide ${currentSlide + 1}`;
      nextTitle.innerText = 'Próximo';
      nextSub.innerText = `Avançar Slide ${currentSlide + 1}`;
    } else {
      nextBadge.classList.add('hidden');
      nextTitle.innerText = 'Próximo';
      nextSub.innerText = 'Fim da Apresentação';
    }

    if (state.next_thumb) {
      nextThumbImg.src = state.next_thumb;
      nextThumbImg.classList.remove('hidden');
    } else {
      nextThumbImg.classList.add('hidden');
    }
  }
}

// Troca de Apresentação disparada pelo Seletor Web
function onPresentationSelected(selectedName) {
  if (!selectedName) return;
  sendCommand('select_presentation', { name: selectedName });
  showToast(`Apresentação selecionada: ${selectedName}`);
}

// Ações dos Botões
function onNextClick() {
  sendCommand('next');
}

function onPrevClick() {
  sendCommand('prev');
}

function onBlackClick() {
  sendCommand('toggle_black');
}

function onWhiteClick() {
  sendCommand('toggle_white');
}

function onStartStopClick() {
  if (presentationState.is_presenting) {
    sendCommand('exit_slideshow');
  } else {
    sendCommand('start_slideshow');
  }
}

// Modal de Salto de Slide
function openJumpModal() {
  const modal = document.getElementById('jumpModal');
  const input = document.getElementById('jumpSlideInput');
  const maxLabel = document.getElementById('jumpMaxSlides');
  
  maxLabel.innerText = `de ${presentationState.total_slides || '?'}`;
  input.value = presentationState.currentSlide || 1;
  input.max = presentationState.total_slides || 999;
  
  modal.classList.remove('hidden');
  input.focus();
}

function closeJumpModal() {
  document.getElementById('jumpModal').classList.add('hidden');
}

function executeJump() {
  const input = document.getElementById('jumpSlideInput');
  const slideNum = parseInt(input.value, 10);
  if (slideNum > 0) {
    sendCommand('goto', { slide: slideNum });
    closeJumpModal();
  }
}

// Modal de Configuração
function toggleConfigModal(forceOpen = false) {
  const modal = document.getElementById('configModal');
  if (forceOpen) {
    modal.classList.remove('hidden');
  } else {
    modal.classList.toggle('hidden');
  }
  document.getElementById('connectionErrorMsg').classList.add('hidden');
}

function saveAndConnect() {
  const roomId = document.getElementById('roomIdInput').value.trim();
  const supabaseUrl = document.getElementById('supabaseUrlInput').value.trim();
  const supabaseKey = document.getElementById('supabaseKeyInput').value.trim();

  if (!roomId || !supabaseUrl || !supabaseKey) {
    showError('Por favor preencha todos os campos.');
    return;
  }

  currentConfig.roomId = roomId;
  currentConfig.supabaseUrl = supabaseUrl;
  currentConfig.supabaseKey = supabaseKey;

  toggleConfigModal(false);
  connectToSupabase();
}

function showError(msg) {
  const errEl = document.getElementById('connectionErrorMsg');
  errEl.innerText = msg;
  errEl.classList.remove('hidden');
}

// Cronômetro
function toggleTimer() {
  const btn = document.getElementById('timerToggleBtn');
  if (timerRunning) {
    clearInterval(timerInterval);
    timerRunning = false;
    btn.innerText = 'Continuar';
  } else {
    timerInterval = setInterval(() => {
      timerSeconds++;
      updateTimerDisplay();
    }, 1000);
    timerRunning = true;
    btn.innerText = 'Pausar';
  }
}

function resetTimer() {
  clearInterval(timerInterval);
  timerRunning = false;
  timerSeconds = 0;
  updateTimerDisplay();
  document.getElementById('timerToggleBtn').innerText = 'Iniciar';
}

function updateTimerDisplay() {
  const mins = Math.floor(timerSeconds / 60);
  const secs = timerSeconds % 60;
  const formatted = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  document.getElementById('timerText').innerText = formatted;
}

// Tamanho da fonte das notas
function changeFontSize(delta) {
  notesFontSize = Math.max(0.75, Math.min(1.8, notesFontSize + delta * 0.15));
  document.getElementById('speakerNotesText').style.fontSize = `${notesFontSize}rem`;
}

// Vibração Haptic
function triggerHaptic() {
  if ('vibrate' in navigator) {
    try {
      navigator.vibrate(35);
    } catch (e) {}
  }
}

// Toast
function showToast(text) {
  const toast = document.getElementById('toast');
  toast.innerText = text;
  toast.classList.remove('hidden');
  setTimeout(() => {
    toast.classList.add('hidden');
  }, 2400);
}

// Atalhos de teclado para teste no PC
function setupKeyboardShortcuts() {
  window.addEventListener('keydown', (e) => {
    // Não ativa atalhos se estiver digitando em um input
    if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;

    if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
      onNextClick();
    } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
      onPrevClick();
    } else if (e.key === 'b' || e.key === 'B') {
      onBlackClick();
    } else if (e.key === 'w' || e.key === 'W') {
      onWhiteClick();
    }
  });
}
