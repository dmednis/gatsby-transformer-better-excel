const fs = require('fs').promises;
const XLSX = require(`xlsx`);

const loadBinaryContent = (fileNode, fallback) => {
    const { absolutePath } = fileNode;
    return absolutePath ? fs.readFile(absolutePath, 'binary') : fallback(fileNode);
};

async function onCreateNode(
    { node, actions, loadNodeContent, createNodeId, createContentDigest },
    options = {}
) {
    const { createNode, createParentChildLink } = actions;
    const extensions = [
        'xls', 'xlsx', 'xlsm', 'xlsb',
        'xml', 'xlw', 'xlc', 'csv',
        'txt', 'dif', 'sylk', 'slk',
        'prn', 'ods', 'fods', 'uos',
        'dbf', 'wks', '123', 'wq1',
        'qpw', 'htm', 'html'
    ];

    if (!extensions.includes((node.extension || '').toLowerCase())) {
        return;
    }

    // Load binary string
    const content = await loadBinaryContent(node, loadNodeContent);

    // accept *all* options to pass to the sheet_to_json function
    const xlsxOptions = {
        ...options,
        // alias legacy `rawOutput` to correct `raw` attribute if raw isn't already defined
        raw: options.raw || options.rawOutput,
        defval: options.defval || options.defaultValue,
    };

    const workbook = XLSX.read(content, { type: `binary`, cellDates: true });
    const workbookContentDigest = createContentDigest(workbook);
    const workbookNode = {
        name: node.name,
        id: createNodeId(node.name),
        children: [],
        parent: node.id,
        internal: {
            contentDigest: workbookContentDigest,
            type: `ExcelWorkbook`,
        },
    };
    createNode(workbookNode);
    createParentChildLink({ parent: node, child: workbookNode });

    workbook.SheetNames.forEach((worksheetName) => {
        const worksheet = workbook.Sheets[worksheetName];
        const parsedRows = XLSX.utils.sheet_to_json(worksheet, xlsxOptions);
        const worksheetContentDigest = createContentDigest(parsedRows);
        const worksheetNode = {
            name: worksheetName,
            id: createNodeId(`${node.name}_${worksheetName}`),
            children: [],
            parent: workbookNode.id,
            internal: {
                contentDigest: worksheetContentDigest,
                type: `ExcelWorksheet`,
            },
        };

        createNode(worksheetNode);
        createParentChildLink({ parent: workbookNode, child: worksheetNode });

        parsedRows.forEach((row, idx) => {
            const rowContentDigest = createContentDigest(row);
            const rowNode = {
                ...row,
                id: createNodeId(`${node.name}_${worksheetName}_${idx}`),
                children: [],
                parent: worksheetNode.id,
                internal: {
                    contentDigest: rowContentDigest,
                    type: 'ExcelWorksheetRow',
                },
            };
            createNode(rowNode);
            createParentChildLink({ parent: worksheetNode, child: rowNode });
        });
    });
}

exports.onCreateNode = onCreateNode;
