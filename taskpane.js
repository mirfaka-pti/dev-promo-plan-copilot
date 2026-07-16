const PROMO_PLAN_KEYWORD = "promo plan";

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

            document.getElementById("errorContainer").innerHTML = `
                <div class="error">
                    <h3>Add-in tidak dapat digunakan</h3>

                    <div>
                        URL:
                    </div>

                    <div>
                        ${escapeHtml(workbookUrl || "(URL tidak tersedia)")}
                    </div>
                </div>
            `;

            return;
        }

        document.getElementById("app").style.display = "block";

        document
            .getElementById("btnCreateLog")
            .addEventListener("click", createLog);

    }
    catch (error) {

        document.getElementById("errorContainer").innerHTML = `
            <div class="error">
                ${error.message}
            </div>
        `;
    }

});

async function createLog() {

    try {

        await Excel.run(async (context) => {

            const worksheet =
                context.workbook.worksheets.getItem("Automation Log");

            const usedRange =
                worksheet.getUsedRange();

            usedRange.load("rowCount");

            await context.sync();

            const nextRow =
                usedRange.rowCount;

            const now = new Date();

            const executionId =
                formatExecutionId(now);

            const executionStartTime =
                formatDisplayDate(now);

            const rowData = [[
                executionId,
                executionStartTime,
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
                    rowData[0].length
                )
                .values = rowData;

            await context.sync();

        });

        document.getElementById("result").innerHTML =
            "<span class='success'>Log berhasil dibuat</span>";

    }
    catch (error) {

        document.getElementById("result").innerHTML =
            `<span style="color:red;">${error.message}</span>`;
    }
}

function formatExecutionId(date) {

    const yy =
        String(date.getFullYear()).slice(-2);

    const mm =
        String(date.getMonth() + 1).padStart(2, "0");

    const dd =
        String(date.getDate()).padStart(2, "0");

    const hh =
        String(date.getHours()).padStart(2, "0");

    const mi =
        String(date.getMinutes()).padStart(2, "0");

    const ss =
        String(date.getSeconds()).padStart(2, "0");

    return `${yy}${mm}${dd}-${hh}${mi}${ss}`;
}

function formatDisplayDate(date) {

    const dd =
        String(date.getDate()).padStart(2, "0");

    const mm =
        String(date.getMonth() + 1).padStart(2, "0");

    const yyyy =
        date.getFullYear();

    const hh =
        String(date.getHours()).padStart(2, "0");

    const mi =
        String(date.getMinutes()).padStart(2, "0");

    const ss =
        String(date.getSeconds()).padStart(2, "0");

    return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

function escapeHtml(text) {

    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}