(async () => {
  await Office.onReady();

  document.getElementById("ok-button").onclick = sendStringToParentPage;
  title_element = document.getElementById("messageTitle");
  message_element = document.getElementById("message");

  const url = new URL(document.URL);
  
  title_element.innerText = url.searchParams.get("messageTitle");
  message_element.innerText = url.searchParams.get("message");

  function sendStringToParentPage() {
    Office.context.ui.messageParent("close dialog");
}
})();
