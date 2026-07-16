const PROMO_PLAN_KEYWORD = "promo";

/* ===========================
   CONFIGURATION
=========================== */

const GET_NAME_URL =
    "https://YOUR_DOMAIN/webhook/get-name";

const START_PROMO_URL =
    "https://YOUR_DOMAIN/webhook/start-promo-plan";

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

        document.getElementById("app")
            .style.display = "block";

        await loadSheetNames();

        document
            .getElementById("btnCreateLog")
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

        const payload = {

            sheetId:
                selectedOption.value,

            sheetName:
                selectedOption.dataset.name,

            brand:
                document.getElementById(
                    "brand"
                ).value,

            promoType:
                document.getElementById(
                    "promoType"
                ).value,

            workbookUrl:
                Office.context.document.url
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

        await insertAutomationLog();

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

async function insertAutomationLog() {

    await Excel.run(
        async (context) => {

            const worksheet =
                context
                    .workbook
                    .worksheets
                    .getItem(
                        "Automation Log"
                    );

            const usedRange =
                worksheet
                    .getUsedRange();

            usedRange.load(
                "rowCount"
            );

            await context.sync();

            const nextRow =
                usedRange.rowCount;

            const now =
                new Date();

            const values = [[

                formatExecutionId(
                    now
                ),

                formatDisplayDate(
                    now
                ),

                "Hello World",

                "Hello world 2",

                "Dummy User",

                "On going"

            ]];

            worksheet
                .getRangeByIndexes(
                    nextRow,
                    0,
                    1,
                    values[0].length
                )
                .values =
                values;

            await context.sync();

        }
    );

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