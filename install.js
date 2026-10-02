(function () {
  "use strict";

  const PLACEMENTS = [
    {
      code: "CRM_DEAL_DETAIL_TAB",
      title: "Excel таблицы в CRM",
      query: "",
    },
    {
      code: "LEFT_MENU",
      title: "Excel таблицы в CRM",
      query: "?mode=crm-menu",
      optional: true,
    },
  ];
  const SHARED_STORAGE_ENTITY = "exctabb24";
  const SHARED_STORAGE_PROPERTY = "DATA";

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

  async function bindPlacement(placement) {
    const handler = new URL(`index.html${placement.query}`, window.location.href).href;

    await callMethod("placement.unbind", {
      PLACEMENT: placement.code,
    }).catch(() => null);

    await callMethod("placement.unbind", {
      PLACEMENT: placement.code,
      HANDLER: handler,
    }).catch(() => null);

    try {
      await callMethod("placement.bind", {
        PLACEMENT: placement.code,
        HANDLER: handler,
        TITLE: placement.title,
      });
      setStatus("Встраивание приложения зарегистрировано.");
    } catch (error) {
      const message = String(error.message || error);
      if (/already|exist|уже/i.test(message)) {
        setStatus("Встраивание приложения уже зарегистрировано.");
        return;
      }
      if (placement.optional) return;
      throw error;
    }
  }

  async function bindPlacements() {
    for (const placement of PLACEMENTS) {
      await bindPlacement(placement);
    }
  }

  async function ensureSharedStorage() {
    await callMethod("entity.add", {
      ENTITY: SHARED_STORAGE_ENTITY,
      NAME: "Excel Tab B24 shared storage",
      ACCESS: { AU: "W" },
    }).catch(() => null);

    await callMethod("entity.update", {
      ENTITY: SHARED_STORAGE_ENTITY,
      ACCESS: { AU: "W" },
    }).catch(() => null);

    await callMethod("entity.item.property.add", {
      ENTITY: SHARED_STORAGE_ENTITY,
      PROPERTY: SHARED_STORAGE_PROPERTY,
      NAME: "Serialized data",
      TYPE: "S",
    }).catch(() => null);
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
        await ensureSharedStorage();
        await bindPlacements();
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
