const PROMO_PLAN_KEYWORD = "promo";

/* ===========================
   CONFIGURATION
=========================== */

const GET_NAME_URL =
    "https://n8n.parainfra.id/webhook/f3ed50fa-9fda-4370-bea5-3cfaa8254c3e";

const START_PROMO_URL =
    "https://n8n.parainfra.id/webhook/1ddb20c4-852a-4060-8aa7-62db6a52da4f";

const BASIC_AUTH =
    "ZGVtb0BkZW1vLmNvbTpEZW1vMTIzNCE=";

/* ===========================
   INIT
=========================== */

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

        document.getElementById("info")
            .style.display = "block";
        document.getElementById("app")
            .style.display = "block";

        await loadSheetNames();

        document
            .getElementById("btnLoadMissingSku")
            .addEventListener(
                "click",
                createLog
            );

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
/* ===========================
   LOAD SHEET NAME
=========================== */

async function loadSheetNames() {

    const response = await fetch(
        GET_NAME_URL,
        {
            method: "GET",
            headers: {
                Authorization:
                    "Basic " + BASIC_AUTH
            }
        }
    );

    if (!response.ok) {

        throw new Error(
            "Gagal mengambil sheet name"
        );
    }

    const data =
        await response.json();

    const dropdown =
        document.getElementById(
            "sheetName"
        );

    dropdown.innerHTML = "";

    data.forEach(item => {

        const option =
            document.createElement(
                "option"
            );

        option.value =
            item.id;

        option.textContent =
            item.name;

        option.dataset.name =
            item.name;

        dropdown.appendChild(
            option
        );

    });

}

/* ===========================
   START PROMO
=========================== */
/* ===========================
   GET ANOMALY SKU CODES
=========================== */

async function getAnomalySkuCodes(tableId) {

    return await Excel.run(async (context) => {
        const tableName = tableId === "{00000000-0001-0000-0000-000000000000}" ? "PromoPlan1" : "PromoPlan2";
        const table =
            context.workbook.tables.getItem(tableName);

        // Range body tabel (tanpa header & total row)
        const bodyRange =
            table.getDataBodyRange();

        const colB =
            bodyRange.getColumn(1); // kolom B (0-based index 1)

        const colC =
            bodyRange.getColumn(2); // kolom C (0-based index 2)

        // Ambil value kolom C dalam 1 batch (murah)
        colC.load("values");

        // Ambil HANYA fill color kolom B dalam 1 batch,
        // bukan getCell() per baris (menghindari 20rb proxy object)
        const cellProps =
            colB.getCellProperties({
                format: {
                    fill: {
                        color: true
                    }
                }
            });

        await context.sync();

        const values = colC.values;         // [[v],[v],...]
        const props = cellProps.value;      // [[{format:{fill:{color}}}],...]

        const TARGET_COLOR = "#B4A7D6";
        const result = [];

        for (let i = 0; i < props.length; i++) {

            const color =
                props[i][0]?.format?.fill?.color;

            if (
                color &&
                color.toUpperCase() === TARGET_COLOR
            ) {
                const code = values[i][0];

                if (code !== "" && code !== null) {
                    result.push(code);
                }
            }
        }

        return result;
    });
}

async function createLog() {

    try {

        const sheetDropdown =
            document.getElementById(
                "sheetName"
            );

        const selectedOption =
            sheetDropdown.options[
                sheetDropdown.selectedIndex
            ];
        
        const now =
            new Date();

        // Ambil kode SKU yang kolom B-nya berwarna #b4a7d6
        const anomalySkuCodes =
            await getAnomalySkuCodes(selectedOption.value);

        const payload = {
            sheetId:
                selectedOption.value,
            program:
                "loadMissingSku",
            execId:
                formatExecutionId(
                    now
                ),
            checkItems: anomalySkuCodes 
        };
        const response =
            await fetch(
                START_PROMO_URL,
                {
                    method: "POST",
                    headers: {
                        "Content-Type":
                            "application/json",

                        Authorization:
                            "Basic " + BASIC_AUTH
                    },
                    body:
                        JSON.stringify(
                            payload
                        )
                }
            );

        if (!response.ok) {

            throw new Error(
                "Gagal menjalankan promo plan"
            );
        }

        await insertAutomationLog(now);

        showToast(
            "Berhasil ditambahkan"
        );

    }
    catch (error) {

        console.error(error);

        showToast(
            error.message
        );
    }

}

/* ===========================
   EXCEL LOG
=========================== */

async function insertAutomationLog(now) {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Automation Log");

    // Ambil tabel berdasarkan nama (pastikan nama tabel di Excel benar)
    const table = worksheet.tables.getItem("AutomationLog");

    // Ambil data dari elemen taskpane
    const sheetDropdown = document.getElementById("sheetName");
    const selectedOption = sheetDropdown.options[sheetDropdown.selectedIndex];
    const sheetId = selectedOption.value;
    const sheetName = selectedOption.dataset.name;

    // Siapkan data untuk baris baru
    const values = [[
      formatExecutionId(now),
      formatDisplayDate(now),
      sheetName,
      `Load Missing SKU`,
      "Dummy User",
      "On going",
      ""
    ]];

    // Tambahkan baris ke tabel
    table.rows.add(null, values);

    await context.sync();
  });
}


/* ===========================
   TOAST
=========================== */

function showToast(message) {

    const toast =
        document.getElementById(
            "toast"
        );

    toast.textContent =
        message;

    toast.classList.add(
        "show"
    );

    setTimeout(() => {

        toast.classList.remove(
            "show"
        );

    }, 3000);

}

/* ===========================
   HELPER
=========================== */

function formatExecutionId(date) {

    const yy =
        String(
            date.getFullYear()
        ).slice(-2);

    const mm =
        String(
            date.getMonth() + 1
        ).padStart(2, "0");

    const dd =
        String(
            date.getDate()
        ).padStart(2, "0");

    const hh =
        String(
            date.getHours()
        ).padStart(2, "0");

    const mi =
        String(
            date.getMinutes()
        ).padStart(2, "0");

    const ss =
        String(
            date.getSeconds()
        ).padStart(2, "0");

    return `${yy}${mm}${dd}-${hh}${mi}${ss}`;
}

function formatDisplayDate(date) {

    const dd =
        String(
            date.getDate()
        ).padStart(2, "0");

    const mm =
        String(
            date.getMonth() + 1
        ).padStart(2, "0");

    const yyyy =
        date.getFullYear();

    const hh =
        String(
            date.getHours()
        ).padStart(2, "0");

    const mi =
        String(
            date.getMinutes()
        ).padStart(2, "0");

    const ss =
        String(
            date.getSeconds()
        ).padStart(2, "0");

    return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

function escapeHtml(text = "") {

    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}