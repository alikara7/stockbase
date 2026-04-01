// theme-switcher.js
// Handles theme switching for Stockbase

const darkTheme = `:root {
  --bg:       #0a0a0a;
  --surface:  #141414;
  --surface2: #1c1c1c;
  --surface3: #242424;
  --border:   #2a2a2a;
  --border2:  #363636;
  --accent:   #e8ff47;
  --blue:     #3b82f6;
  --blue-d:   rgba(59,130,246,.15);
  --blue-b:   rgba(59,130,246,.3);
  --green:    #22c55e;
  --green-d:  rgba(34,197,94,.15);
  --green-b:  rgba(34,197,94,.3);
  --purple:   #a78bfa;
  --purple-d: rgba(167,139,250,.15);
  --purple-b: rgba(167,139,250,.3);
  --danger:   #f87171;
  --danger-d: rgba(248,113,113,.12);
  --danger-b: rgba(248,113,113,.3);
  --warn:     #fbbf24;
  --warn-d:   rgba(251,191,36,.12);
  --text:     #f1f5f9;
  --text2:    #94a3b8;
  --muted:    #4b5563;
  --shadow:   0 1px 3px rgba(0,0,0,.4);
  --shadow-md:0 4px 12px rgba(0,0,0,.5);
  --shadow-lg:0 20px 40px rgba(0,0,0,.6);
}`;

const lightTheme = `:root {
  --bg:       #f0f4f8;
  --surface:  #ffffff;
  --surface2: #f8fafc;
  --border:   #e2e8f0;
  --border2:  #cbd5e1;
  --green:    #16a34a;
  --green-l:  #dcfce7;
  --green-m:  #86efac;
  --blue:     #2563eb;
  --blue-l:   #dbeafe;
  --blue-m:   #93c5fd;
  --danger:   #dc2626;
  --danger-l: #fee2e2;
  --warn:     #d97706;
  --warn-l:   #fef3c7;
  --text:     #0f172a;
  --text2:    #475569;
  --muted:    #94a3b8;
  --shadow:   0 1px 3px rgba(0,0,0,.08), 0 1px 2px rgba(0,0,0,.04);
  --shadow-md:0 4px 6px rgba(0,0,0,.07), 0 2px 4px rgba(0,0,0,.04);
  --shadow-lg:0 10px 24px rgba(0,0,0,.1), 0 4px 8px rgba(0,0,0,.05);
}`;

function setTheme(theme) {
  document.getElementById('theme-style').textContent = theme === 'dark' ? darkTheme : lightTheme;
  document.getElementById('theme-toggle').textContent = theme === 'dark' ? '🌙' : '☀️';
  localStorage.setItem('stockbase-theme', theme);
}

function toggleTheme() {
  const current = localStorage.getItem('stockbase-theme') || 'dark';
  setTheme(current === 'dark' ? 'light' : 'dark');
}

// On load, set theme from localStorage or system preference
(function() {
  let theme = localStorage.getItem('stockbase-theme');
  if (!theme) {
    theme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  setTheme(theme);
  document.getElementById('theme-toggle').onclick = toggleTheme;
})();
