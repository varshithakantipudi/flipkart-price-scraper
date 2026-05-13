document.addEventListener('DOMContentLoaded', function () {
    const startBtn = document.getElementById('startBtn');
    const fileInput = document.getElementById('fileInput');
    const status = document.getElementById('status');

    startBtn.addEventListener('click', async () => {
        if (!fileInput.files.length) {
            alert("Please upload an Excel file first!");
            return;
        }

        try {
            status.innerText = "Reading file...";
            const file = fileInput.files[0];
            const data = await file.arrayBuffer();
            
            // Read Excel
            const workbook = XLSX.read(data);
            const sheetName = workbook.SheetNames[0];
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

            // Get FSNs
            const fsnList = json.map(row => row.FSN).filter(fsn => fsn);
            
            if (fsnList.length === 0) {
                status.innerText = "Error: No 'FSN' column found.";
                return;
            }

            const results = [];
            status.innerText = `Scraping 0/${fsnList.length}...`;

            for (let i = 0; i < fsnList.length; i++) {
                const fsn = fsnList[i];
                try {
                    const response = await fetch(`https://www.flipkart.com/product/p/itme?pid=${fsn}`);
                    const text = await response.text();
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(text, 'text/html');

                    // Price Scraper Logic
                    const priceElement = doc.querySelector("div.v1zwn21l.v1zwn20._1psv1zeb9._1psv1ze0");
                    const price = priceElement ? priceElement.innerText.trim() : "N/A";

                    results.push({ FSN: fsn, Price: price });
                    status.innerText = `Progress: ${i + 1}/${fsnList.length}`;
                } catch (err) {
                    results.push({ FSN: fsn, Price: "Error" });
                }
                
                // Small delay to prevent browser hang
                await new Promise(r => setTimeout(r, 100));
            }

            // Save File
            const newSheet = XLSX.utils.json_to_sheet(results);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Prices");
            XLSX.writeFile(newWorkbook, "flipkart_prices.xlsx");
            
            status.innerText = "Done! Downloaded.";

        } catch (error) {
            console.error(error);
            status.innerText = "Error reading file.";
        }
    });
});