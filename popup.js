document.getElementById("v").textContent = chrome.runtime.getManifest().version;
chrome.runtime.sendMessage({ type: "fetch", path: "/php/ajax/drivertracking/get.driver.table.php" }, (r) => {
  const dot = document.getElementById("nd"), txt = document.getElementById("nt");
  if (r && r.ok) { dot.className = "dot ok"; txt.textContent = "Logged in to Nexl"; }
  else {
    dot.className = "dot bad";
    txt.innerHTML = 'Not logged in. <a href="https://controller.nexl.online/" target="_blank">Open Nexl</a>';
  }
});
