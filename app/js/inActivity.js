const AUTO_LOGOUT_TIME = 15 * 60 * 1000;
let inactivityTimer = null;

function startInactivityTimer() {
  clearTimeout(inactivityTimer);
  inactivityTimer = setTimeout(() => {
    alert("Logged out due to inactivity");
    logout();
  }, AUTO_LOGOUT_TIME);
}

function resetInactivityTimer() {
  if (localStorage.getItem("loggedIn") === "true") {
    startInactivityTimer();
  }
}

["click", "mousemove", "keydown", "scroll", "touchstart"].forEach(event => {
  document.addEventListener(event, resetInactivityTimer, true);
});