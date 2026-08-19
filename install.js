(function () {
  "use strict";

  const PLACEMENT_CODE = "CRM_DEAL_DETAIL_TAB";
  const PLACEMENT_TITLE = "Excel Tab B24";

  function callMethod(method, params = {}) {
    return new Promise((resolve, reject) => {
      if (!window.BX24 || typeof window.BX24.callMethod !== "function") {
        reject(new Error("Bitrix24 SDK is not available"));
        return;
      }

      window.BX24.callMethod(method, params, (result) => {
        if (result.error()) {
          reject(new Error(result.error_description() || result.error()));
          return;
        }

        resolve(result.data());
      });
    });
  }

  function setStatus(message) {
    const node = document.getElementById("installStatus");
    if (node) node.textContent = message;
  }

  async function bindPlacement() {
    const handler = new URL("index.html", window.location.href).href;

    await callMethod("placement.unbind", {
      PLACEMENT: PLACEMENT_CODE,
    }).catch(() => null);

    await callMethod("placement.unbind", {
      PLACEMENT: PLACEMENT_CODE,
      HANDLER: handler,
    }).catch(() => null);

    try {
      await callMethod("placement.bind", {
        PLACEMENT: PLACEMENT_CODE,
        HANDLER: handler,
        TITLE: PLACEMENT_TITLE,
      });
      setStatus("Вкладка сделки зарегистрирована.");
    } catch (error) {
      const message = String(error.message || error);
      if (/already|exist|уже/i.test(message)) {
        setStatus("Вкладка сделки уже зарегистрирована.");
        return;
      }
      throw error;
    }
  }

  function finishInstall() {
    if (window.BX24 && typeof window.BX24.installFinish === "function") {
      window.BX24.installFinish();
      return;
    }

    setStatus("Установка готова. Закройте это окно в Bitrix24.");
  }

  function boot() {
    const finishButton = document.getElementById("finishButton");
    if (finishButton) finishButton.addEventListener("click", finishInstall);

    if (!window.BX24 || typeof window.BX24.init !== "function") {
      setStatus("Откройте установку внутри Bitrix24.");
      return;
    }

    window.BX24.init(async () => {
      try {
        setStatus("Регистрируем вкладку в карточке сделки...");
        await bindPlacement();
        if (finishButton) finishButton.disabled = false;
        finishInstall();
      } catch (error) {
        setStatus(`Ошибка установки: ${error.message || error}`);
        if (finishButton) finishButton.disabled = false;
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
