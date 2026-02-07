
const LogsURL = '/api/logs/activity';

async function fetchLogs() {
    const terminal = document.getElementById('logTerminal');
    // Don't clear local logs immediately, maybe append? 
    // Actually, for "System Activity" from server, we probably want to show the list.
    // The current UI concept in dashboard.html is a terminal.
    // Let's prepend server logs or just replace.
    // For now, let's keep the terminal for "current session" logs, and maybe add a section for "History" or just fetch and fill.

    // User requested "save older logs locally too" and "detailed error logging".
    // I will fetch the server logs and display them formatted in the terminal style.

    try {
        const res = await fetch(LogsURL);
        const logs = await res.json();

        if (logs.length === 0) return;

        // We want to merge or show them.
        // Let's clear and show recent 100 on refresh
        terminal.innerHTML = '';

        // Reverse to show oldest first if we append, but logs API returns newest first.
        // Terminal usually puts newest at bottom.
        const reversed = logs.reverse();

        reversed.forEach(l => {
            const div = document.createElement('div');
            let typeClass = 'info';
            if (l.type.includes('ERROR')) typeClass = 'error';
            if (l.type.includes('SENT')) typeClass = 'success';

            const time = new Date(l.createdAt).toLocaleTimeString();
            div.className = `log-line ${typeClass}`;
            div.textContent = `[${time}] [${l.type}] ${l.description}`;
            terminal.appendChild(div);
        });

        // Scroll to bottom
        terminal.scrollTop = terminal.scrollHeight;

    } catch (e) {
        console.error("Failed to fetch logs", e);
    }
}

document.getElementById('refreshLogsBtn').addEventListener('click', fetchLogs);
