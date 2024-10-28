document.getElementById('extractButton').addEventListener('click', extractComments);

document.getElementById('fileInput').addEventListener('change', function(event) {
    var fileCount = event.target.files.length;
    var fileCountText = fileCount > 0 ? fileCount + ' file(s) chosen' : 'No files chosen';
    document.getElementById('fileCount').textContent = fileCountText;

    // Set the upload date when files are selected
    if (fileCount > 0) {
        const uploadDate = formatDate(new Date()); // Format as DD.MM.YYYY
        document.getElementById('uploadDate').value = uploadDate;
    }
});

async function extractComments() {
    const fileInput = document.getElementById('fileInput').files;
    if (fileInput.length === 0) {
        alert('Please upload at least one PDF file.');
        return;
    }

    const uploadDate = document.getElementById('uploadDate').value || formatDate(new Date());

    let comments = [];
    let uniqueComments = new Set(); // Set to track unique comments

    for (const file of fileInput) {
        const pdfData = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;

        const fileName = file.name;
        const weekMatch = fileName.match(/(WK\d+)/);
        const week = weekMatch ? `Week-${weekMatch[0].substring(2)}` : 'Unknown';
        const bannerName = 'Walmart';

        const pageMatch = fileName.match(/^WK\d+_24_(.+?)(?:(_CPR|_PI\s*\d*|_CF|_PR\s*\d*|_PR[1-4](_CF)?|\.pdf))?$/);
        let page = pageMatch ? pageMatch[1] : 'Unknown';
        page = page.replace(/(_CPR|_PI\s*\d*|_CF|_PR\s*\d*|_PR[1-4](_CF)?|_CF)?$/, '').trim();

        let proofName = 'Press';
        if (fileName.includes('_CPR')) {
            proofName = 'CPR';
        } else if (/(_PR\d)/.test(fileName)) {
            const prMatch = fileName.match(/_PR(\d)/);
            proofName = `Proof ${prMatch[1]}`;
        }

        let biEngAll = 'All Zones';
        const bilZonesPatterns = ['B_ON', 'B_NB', 'B_QC', '_B_MTL_RADDAR'];
        const engZonesPatterns = ['E_ON', 'NB', 'NS', 'PE', 'NL', 'MB', 'SK', 'AB', 'BC', 'E_NAT', 'E_ATL', 'E_WEST', '_E_VAN_RADDAR'];

        if (bilZonesPatterns.some(pattern => fileName.includes(pattern))) {
            biEngAll = 'Bil Zones';
        } else if (engZonesPatterns.some(pattern => fileName.includes(pattern))) {
            biEngAll = 'Eng Zones';
        }

        for (let i = 0; i < pdf.numPages; i++) {
            const pdfPage = await pdf.getPage(i + 1);
            const annotations = await pdfPage.getAnnotations();

            annotations.forEach(annotation => {
                if (annotation.subtype !== 'Popup') {
                    let errorType = 'Product_Description';
                    let content = annotation.contents || 'No content';
                    
                    // Normalize multi-line content to a single line
                    content = content.replace(/\r?\n|\r/g, ' ').trim();

                    if (content.toLowerCase().includes('price')) {
                        errorType = 'Price_Point';
                    } else if (content.toLowerCase().includes('alignment')) {
                        errorType = 'Overall_Layout';
                    } else if (content.toLowerCase().includes('image')) {
                        errorType = 'Image_Usage';
                    }

                    let errorsContent = '';
                    let gdContent = '';

                    if (content.startsWith('GD:')) {
                        gdContent = content.substring(3).trim();
                    } else {
                        errorsContent = content;
                    }

                    const commentKey = `${page}_${proofName}_${errorsContent}`; // Unique key for duplicates

                    if (!uniqueComments.has(commentKey)) {
                        uniqueComments.add(commentKey); // Add to set if unique
                        
                        comments.push({
                            Date: uploadDate,
                            Banner: bannerName,
                            Week: week,
                            Page: page,
                            Proof: proofName,
                            BI_ENG_All: biEngAll,
                            PageAssembler: '', // Placeholder
                            QC: annotation.title || 'Unknown',
                            SJC_QC: '', // Placeholder
                            Correction_Revision: gdContent,
                            NO_OF_ERRORS: 1,
                            ERROR_CATEGORY: errorType,
                            REMARKS: errorsContent
                        });
                    }
                }
            });
        }

        comments.push({
            Date: '',
            Banner: '',
            Week: '',
            Page: '',
            Proof: '',
            BI_ENG_All: '',
            PageAssembler: '',
            QC: '',
            SJC_QC: '',
            Correction_Revision: '',
            NO_OF_ERRORS: '',
            ERROR_CATEGORY: '',
            REMARKS: ''
        });
    }

    exportToExcel(comments);
}

function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`; // Format as DD.MM.YYYY
}

function exportToExcel(comments) {
    const worksheet = XLSX.utils.json_to_sheet(comments, { header: ["Date", "Banner", "Week", "Page", "Proof", "BI_ENG_All", "PageAssembler", "QC", "SJC_QC", "Correction_Revision", "NO_OF_ERRORS", "ERROR_CATEGORY", "REMARKS"] });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Comments");
    XLSX.writeFile(workbook, "comments.xlsx");
}