document.getElementById('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const staffId = document.getElementById('staffId').value;
    const btn = document.getElementById('loginBtn');
    const errorMsg = document.getElementById('errorMsg');

    // UI Loading state
    btn.innerHTML = 'Validating...';
    btn.disabled = true;
    errorMsg.classList.remove('visible');

    try {
        const res = await fetch('/api/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ staffId })
        });

        const data = await res.json();

        if (res.ok && data.success) {
            // Store token and NAME
            sessionStorage.setItem('iqac_token', data.token);
            sessionStorage.setItem('iqac_staff_name', data.name || 'Staff Member');
            window.location.href = '/dashboard.html';
        } else {
            errorMsg.textContent = data.error || 'Access Denied';
            errorMsg.classList.add('visible');
            btn.innerHTML = 'Enter System';
            btn.disabled = false;
        }
    } catch (err) {
        console.error("Full Login Error:", err);
        // Try to show more detail if available
        let msg = 'Connection Error';
        if (err.message) msg += `: ${err.message}`;
        errorMsg.textContent = msg;
        errorMsg.classList.add('visible');
        btn.innerHTML = 'Enter System';
        btn.disabled = false;
    }
});
