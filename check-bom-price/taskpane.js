(function () {
  "use strict";

  // ===== CONFIGURATION =====
  var WEBHOOK_URL = "https://n8n.parainfra.id/webhook/a6972c9b-c8bd-4f7a-a498-4926e7de3091"; // Replace with your actual webhook URL
  var AUTH_TOKEN = "Basic cGFyYWdvbmR1bW15QGdtYWlsLmNvbTpQYXJhZ29uSmF5YTEyMyE="; // Replace with your actual Basic Auth token
  var PROMO_PLAN_KEYWORD = "promo"; // Only allow this add-in on workbooks whose URL contains this keyword

  // ===== OFFICE INITIALIZATION =====
  // Office.onReady(function (info) {
  //   try {
  //     var isExcel = info.host === Office.HostType.Excel;
  //     var workbookUrl =
  //       (Office && Office.context && Office.context.document && Office.context.document.url) || "";
  //     var isPromoPlan = workbookUrl.toLowerCase().indexOf(PROMO_PLAN_KEYWORD) !== -1;

  //     if (!isExcel || !isPromoPlan) {
  //       showUnavailable(workbookUrl);
  //       return;
  //     }

  //     initTaskpane();
  //   } catch (error) {
  //     console.error(error);
  //     showUnavailable("", error.message);
  //   }
  // });
Office.onReady(async (info) => {

    try {

        const isExcel =
            info.host === Office.HostType.Excel;

        const workbookUrl =
            Office?.context?.document?.url || "";

        const isPromoPlan =
            workbookUrl
                .toLowerCase()
                .includes(PROMO_PLAN_KEYWORD);

        if (!isExcel || !isPromoPlan) {

            document.getElementById(
                "errorContainer"
            ).innerHTML = `
                <div class="error">
                    <h3>Add-in tidak dapat digunakan</h3>
                    <div><b>Workbook URL:</b></div>
                    <div>${escapeHtml(workbookUrl)}</div>
                </div>
            `;

            return;
        }

        document.getElementById("app")
            .style.display = "block";

      initTaskpane();

    }
    catch (error) {

        console.error(error);

        document.getElementById(
            "errorContainer"
        ).innerHTML = `
            <div class="error">
                ${error.message}
            </div>
        `;
    }

});
  // ===== URL / HOST GATE =====
  // Shows the existing "error" state repurposed as a blocking message when
  // the add-in is opened outside Excel, or in an Excel workbook whose URL
  // does not contain PROMO_PLAN_KEYWORD.
  function showUnavailable(workbookUrl, errorMessage) {
    var titleEl = document.querySelector("#state-error .state-title");
    var descEl = document.getElementById("errorDesc");

    if (titleEl) titleEl.textContent = "Add-in tidak dapat digunakan";

    if (descEl) {
      if (errorMessage) {
        descEl.innerHTML = escHtml(errorMessage);
      } else {
        descEl.innerHTML =
          "<b>Workbook URL:</b><br/>" + escHtml(workbookUrl || "(tidak tersedia)");
      }
    }

    showState("error");
  }

  function initTaskpane() {
    var searchInput = document.getElementById("searchInput");
    var searchBtn = document.getElementById("searchBtn");

    searchBtn.addEventListener("click", triggerSearch);

    searchInput.addEventListener("keydown", function (e) {
      if (e.key === "Enter") {
        triggerSearch();
      }
    });
  }

  // ===== STATE MANAGEMENT =====
  var currentState = "initial";

  function showState(name) {
    var states = ["initial", "loading", "success", "empty", "error"];
    states.forEach(function (s) {
      var el = document.getElementById("state-" + s);
      if (el) {
        el.classList.remove("active");
      }
    });
    var target = document.getElementById("state-" + name);
    if (target) target.classList.add("active");
    currentState = name;
  }

  // ===== SEARCH TRIGGER =====
  function triggerSearch() {
    var query = document.getElementById("searchInput").value.trim();
    if (!query) {
      showState("initial");
      return;
    }
    performSearch(query);
  }

  // ===== PERFORM SEARCH (AJAX to Webhook) =====
  function performSearch(sku) {
    showState("loading");

    fetch(WEBHOOK_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: AUTH_TOKEN
      },
      body: JSON.stringify({ sku: sku })
    })
      .then(function (res) {
        if (res.status === 401) throw new Error("Unauthorized");
        if (res.status === 403) throw new Error("Forbidden");
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
      })
      .then(function (data) {
        handleResponse(data, sku);
      })
      .catch(function (err) {
        showState("error");
        var msg = "Gagal terhubung ke server. Silakan coba lagi.";
        if (err.message === "Unauthorized") msg = "Autentikasi gagal. Token tidak valid.";
        if (err.message === "Forbidden") msg = "Akses ditolak. Anda tidak memiliki izin.";
        document.getElementById("errorDesc").textContent = msg;
        console.error("Search error:", err);
      });
  }

  // ===== RESPONSE HANDLER =====
  function handleResponse(data, sku) {
    if (typeof data === "string") data = JSON.parse(data);
    if (!data || data.status !== "success") {
      showState("error");
      document.getElementById("errorDesc").textContent =
        data && data.message ? data.message : "Terjadi kesalahan. Coba lagi.";
      return;
    }

    var items = data.data;

    if (!items || items.length === 0) {
      showState("empty");
      document.getElementById("emptyDesc").textContent =
        'SKU "' + sku + '" tidak ditemukan dalam sistem.';
      return;
    }

    renderResults(items, sku);
  }

  // ===== RENDER RESULTS =====
  function renderResults(items, sku) {
    var label = document.getElementById("resultLabel");
    var container = document.getElementById("resultContainer");

    label.textContent = "SEARCH RESULT FOR '" + sku.toUpperCase() + "'";
    container.innerHTML = "";

    // Render main product (first item as header card)
    var main = items[0];
    var headerEl = document.createElement("div");
    headerEl.className = "product-header";
    headerEl.innerHTML =
      '<div class="product-thumb">' +
      (main.image
        ? '<img src="https://d1wrygdlbvzcxl.cloudfront.net/' + escHtml(main.image) + '" alt=""/>'
        : "<span>" + (main.size || "64px") + "</span>") +
      "</div>" +
      '<div class="product-info">' +
      '<div class="product-name">' + escHtml(main.name) + "</div>" +
      (main.variant ? '<div class="product-sub">' + escHtml(main.variant || "") + "</div>" : "") +
      "</div>";
    container.appendChild(headerEl);

    // Jika hanya 1 item, tampilkan bom-footer di luar product-header
    if (items.length >= 1) {
      var footerEl = document.createElement("div");
      footerEl.className = "bom-footer";
      footerEl.style.cssText = "margin-top: 4px;";
      footerEl.innerHTML =
        '<div class="bom-qty">Harga SAP: <br /><span>' +
        escHtml(String(main.sapPrice || 0)) +
        "</span></div>" +
        '<div class="bom-stock">Harga Satuan: <br /><span>' +
        escHtml(String(main.price || 0)) +
        "</span></div>";
      container.appendChild(footerEl);
    }

    // BOM section (remaining items)
    if (items.length > 1) {
      var sectionLabel = document.createElement("div");
      sectionLabel.className = "section-label";
      sectionLabel.textContent = "DETAIL BOM";
      container.appendChild(sectionLabel);

      items.slice(1).forEach(function (item, idx) {
        var itemEl = document.createElement("div");
        itemEl.className = "bom-item";
        itemEl.style.animationDelay = idx * 0.04 + "s";
        var qtyNum = item.qty !== undefined && item.qty !== null ? item.qty : 0;
        var isCustomGwp = item.isCustomGwp === true || item.isCustomGwp === "true";

        itemEl.innerHTML =
          '<div class="bom-top">' +
          '<div class="bom-thumb-wrap">' +
          '<div class="bom-thumb">' +
          (item.image
            ? '<img src="https://d1wrygdlbvzcxl.cloudfront.net/' +
              escHtml(item.image) +
              '" alt="" style="width:100%;height:100%;object-fit:cover;"/>'
            : "<span>" + (item.size || "48px") + "</span>") +
          "</div>" +
          '<span class="qty-badge">x' + escHtml(String(qtyNum)) + "</span>" +
          "</div>" +
          '<div class="bom-info">' +
          '<div class="bom-info-header">' +
          '<div class="bom-size">' + escHtml(item.sku || "") + "</div>" +
          (isCustomGwp ? '<span class="new-badge">GWP</span>' : "") +
          "</div>" +
          '<div class="bom-desc">' + escHtml(item.name || "") + "</div>" +
          "</div>" +
          "</div>" +
          '<div class="bom-footer">' +
          '<div class="bom-qty">Harga SAP: <br /><span>' +
          escHtml(String(item.sapPrice || 0)) +
          "</span></div>" +
          '<div class="bom-stock">Harga Satuan: <br /><span>' +
          escHtml(String(item.price || 0)) +
          "</span></div>" +
          "</div>";
        container.appendChild(itemEl);
      });
    }

    showState("success");
  }

  // ===== UTILITY =====
  function escHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
})();