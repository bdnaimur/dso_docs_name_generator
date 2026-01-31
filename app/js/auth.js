

// async function hashPassword(password) {
//   const encoder = new TextEncoder();
//   const data = encoder.encode(password);

//   const hashBuffer = await crypto.subtle.digest("SHA-256", data);
//   const hashArray = Array.from(new Uint8Array(hashBuffer));

//   return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
// }

// async function setupUser() {
// const AUTH_CONFIG = {
//   username: "admin",
//   passwordHash: "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"
// };

//   const passwordHash = await hashPassword(AUTH_CONFIG?.);

//   localStorage.setItem("auth_user", JSON.stringify({
//     username,
//     passwordHash
//   }));
// }

// setupUser();


// document.getElementById("loginForm").addEventListener("submit", (e) => {
//   e.preventDefault();
//   login();
// });
// async function login() {
//   const userInput = document.getElementById("username").value.trim();
//   const passInput = document.getElementById("password").value.trim();

// //   if (userInput === username && passInput === passwordHash) {
// //     localStorage.setItem("loggedIn", "true");
// //     updateUser(true);
// //     // startInactivityTimer();
// //   } else {
// //     alert("Incorrect username or password");
// //   }

// const storedUser = JSON.parse(localStorage.getItem("auth_user"));
//   if (!storedUser) return false;

//   const inputHash = await hashPassword(passInput);

//   return (
//     inputUsername === storedUser.username &&
//     inputHash === storedUser.passwordHash
//   );
// }

// function logout() {
//   localStorage.removeItem("loggedIn");
// //   clearTimeout(inactivityTimer);
//   updateUser(false);
// }

// function updateUser(isLoggedIn) {
//   //   const status = document.getElementById("userStatus");
//   const loginForm = document.getElementById("loginForm");
//   const logoutBtn = document.getElementById("logoutBtn");
//   const card = document.querySelector(".card");
//   const authBar = document.getElementById("authBar");

//   if (isLoggedIn) {
//     // status.textContent = `Signed in as: ${username}`;
//     loginForm.style.display = "none";
//     logoutBtn.style.display = "inline-block";
//     card.style.display = "block";
//     authBar.style.display = "none";
//   } else {
//     // status.textContent = "Not signed in";
//     loginForm.style.display = "block";
//     logoutBtn.style.display = "none";
//     card.style.display = "none";
//     authBar.style.display = "flex";
//   }
// }


const AUTH_CONFIG = {
  username: "admin",
  passwordHash: "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"
};

async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);

  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));

  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

function setupUser() {
  if (!localStorage.getItem("auth_user")) {
    localStorage.setItem(
      "auth_user",
      JSON.stringify(AUTH_CONFIG)
    );
  }
}

setupUser();

document.getElementById("loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  await login();
});

async function login() {
  const userInput = document.getElementById("username").value.trim();
  const passInput = document.getElementById("password").value.trim();

  const storedUser = JSON.parse(localStorage.getItem("auth_user"));
  if (!storedUser) {
    alert("User not set up");
    return;
  }

  const inputHash = await hashPassword(passInput);

  if (
    userInput === storedUser.username &&
    inputHash === storedUser.passwordHash
  ) {
    localStorage.setItem("loggedIn", "true");
    updateUser(true);
  } else {
    alert("Incorrect username or password");
  }
}

function logout() {
  localStorage.removeItem("loggedIn");
  updateUser(false);
}

function updateUser(isLoggedIn) {
  const loginForm = document.getElementById("loginForm");
  const logoutBtn = document.getElementById("logoutBtn");
  const card = document.querySelector(".card");
  const authBar = document.getElementById("authBar");

  if (isLoggedIn) {
    loginForm.style.display = "none";
    logoutBtn.style.display = "inline-block";
    card.style.display = "block";
    authBar.style.display = "none";
  } else {
    loginForm.style.display = "block";
    logoutBtn.style.display = "none";
    card.style.display = "none";
    authBar.style.display = "flex";
  }
}

window.addEventListener("load", () => {
  const isLoggedIn = localStorage.getItem("loggedIn") === "true";
  updateUser(isLoggedIn);
});