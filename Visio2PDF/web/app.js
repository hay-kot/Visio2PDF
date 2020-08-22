// Get Folders from User Input... Tkinter Callback in Python
let dir_path;
let cover_path;
let insert_version_tag;

function toggleVersioning() {
  var current_val = document.getElementById("version_tag").disabled;

  if (current_val === true) {
    document.getElementById("version_tag").disabled = false;
    document.getElementById("engineer_name").disabled = false;
    document.getElementById("version_tag").required = true;
    document.getElementById("engineer_name").required = true;
    document.getElementById("engineer_name").placeholder =
      "Required: John C. Doe";
    document.getElementById("version_tag").placeholder = "Required: v1.01";
  } else {
    document.getElementById("version_tag").disabled = true;
    document.getElementById("engineer_name").disabled = true;
    document.getElementById("version_tag").required = false;
    document.getElementById("engineer_name").required = false;
    document.getElementById("version_tag").placeholder = "";
    document.getElementById("engineer_name").placeholder = "";
  }
}

eel.expose(setMsgVisible);
function setMsgVisible() {
  document.getElementById("msg").style.display = "block";
  document.getElementById("working_indicator").style.display = "none";
}

function showLog() {
  var currentState = document.getElementById("output").style.display;

  if (currentState === "none") {
    document.getElementById("output").style.display = "block";
  } else {
    document.getElementById("output").style.display = "none";
  }
}

async function getFolder() {
  dir_path = await eel.btn_SelectDir()();
  let dir_div = document.getElementById("dir_name");
  dir_div.value = dir_path;

  // Calls attempts to find a cover sheet
  cover_path = await eel.find_cover(dir_path)();
  if (cover_path !== false) {
    let cover_div = document.getElementById("cover_name");
    cover_div.value = cover_path;
  }
}

async function getCover() {
  cover_path = await eel.btn_SelectCoverSheet()();
  let cover_div = document.getElementById("cover_name");
  cover_div.value = cover_path;
}

function validateEntry(entry_array) {
  var falseCheck;
  entry_array.forEach(function (item, index, array) {
    if (item === "") {
      console.log("Value is False");
      falseCheck = false;
    } else if (item === "undefined") {
      falseCheck = false;
    }
  });
  console.log(falseCheck);
  if (falseCheck === false) {
    return false;
  } else {
    return true;
  }
}

function validateForm() {
  var sys_name = document.getElementById("sys_name").value;
  var insert_version_tag = document.getElementById("enable_tagging").checked;
  var version_tag = document.getElementById("version_tag").value;
  var engineer_name = document.getElementById("engineer_name").valued;

  if (insert_version_tag === true) {
    var validateArray = [dir_path, sys_name, version_tag, engineer_name];
  } else {
    var validateArray = [dir_path, sys_name];
  }
  console.log(validateArray);

  var isValid = validateEntry(validateArray);
  console.log(isValid);
  return isValid;
}

async function convertVisio() {
  if (validateForm() === false) {
    document.getElementById("required_fields_alert").style.display = "block";
    return;
  }
  document.getElementById("required_fields_alert").style.display = "none";
  document.getElementById("msg").style.display = "none";

  let sys_name = document.getElementById("sys_name").value;
  let version_tag = document.getElementById("version_tag").value;
  let engineer_name = document.getElementById("engineer_name").value;

  let insert_version_tag = document.getElementById("enable_tagging").checked;

  eel.main(
    dir_path,
    cover_path,
    sys_name,
    insert_version_tag,
    version_tag,
    engineer_name
  );
  document.getElementById("working_indicator").style.display = "block";
}

eel.expose(putMessageInOutput);
function putMessageInOutput(message) {
  const outputNode = document.querySelector("#output textarea");
  outputNode.value += message; // Add the message
  if (!message.endsWith("\n")) {
    outputNode.value += "\n"; // If there was no new line, add one
  }

  // Set the correct height to fit all the output and then scroll to the bottom
  // outputNode.style.height = 'auto';
  outputNode.style.height = outputNode.scrollHeight + 10 + "px";
  window.scrollTo(0, document.body.scrollHeight);
}

eel.expose(signalPackagingComplete);
function signalPackagingComplete(successful) {
  setPackagingComplete(successful);
  window.scrollTo(0, document.body.scrollHeight);
}
