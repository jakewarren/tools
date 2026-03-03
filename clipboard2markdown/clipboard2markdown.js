(function () {
    'use strict';

    // ---------------------------------------------------------------------------
    // Turndown (used for non-spreadsheet HTML)
    // ---------------------------------------------------------------------------

    var turndownService = null;

    function getTurndown() {
        if (turndownService) return turndownService;

        turndownService = new TurndownService({
            headingStyle: 'atx',
            hr: '---',
            bulletListMarker: '-',
            codeBlockStyle: 'fenced',
            emDelimiter: '_',
        });

        turndownService.use(turndownPluginGfm.gfm);
        turndownService.remove(['script', 'style', 'head', 'nav', 'footer']);

        return turndownService;
    }

    // ---------------------------------------------------------------------------
    // Direct table → markdown conversion
    //
    // For HTML that includes tables, we convert table DOM directly using
    // textContent to avoid noisy inline styles/wrappers (common in spreadsheet
    // sources) and preserve clean GFM table output.
    // ---------------------------------------------------------------------------

    // Return the plain-text content of a cell, escaping pipe characters so they
    // don't break the GFM table syntax.
    function cellText(cell) {
        return cell.textContent
            .trim()
            .replace(/\s+/g, ' ')       // collapse internal whitespace / newlines
            .replace(/\|/g, '\\|');     // escape pipes
    }

    // Returns true if a row's cells are all (or majority) bold — used as a
    // proxy for "this is a header row" in Google Sheets, which uses <td> for
    // everything but applies font-weight:700 to cells the user has bolded.
    function rowIsBold(row) {
        var cells = Array.from(row.children).filter(function (child) {
            var tag = child.tagName && child.tagName.toLowerCase();
            return tag === 'td' || tag === 'th';
        });
        if (!cells.length) return false;
        var boldCount = cells.filter(function (cell) {
            var style = cell.getAttribute('style') || '';
            return /font-weight\s*:\s*(bold|[6-9]\d\d)/i.test(style) ||
                cell.querySelector('b, strong') !== null;
        }).length;
        // Require a majority of cells to be bold to avoid false positives from
        // a single bold value in an otherwise plain data row.
        return boldCount > cells.length / 2;
    }

    // Convert a single <table> DOM element to a GFM markdown table string.
    // colspan is handled by repeating the cell value; rowspan is ignored (GFM
    // tables have no equivalent).
    function getTableRows(table) {
        var rows = [];
        Array.from(table.children).forEach(function (child) {
            var tag = child.tagName && child.tagName.toLowerCase();
            if (tag === 'tr') {
                rows.push(child);
                return;
            }
            if (tag === 'thead' || tag === 'tbody' || tag === 'tfoot') {
                Array.from(child.children).forEach(function (sectionChild) {
                    var sectionTag = sectionChild.tagName && sectionChild.tagName.toLowerCase();
                    if (sectionTag === 'tr') rows.push(sectionChild);
                });
            }
        });
        return rows;
    }

    function tableToMarkdown(table) {
        var rows = getTableRows(table);
        if (!rows.length) return '';

        // Detect whether the table has a semantic header.
        // - <thead> or <th>: Excel and proper HTML tables
        // - bold first row: Google Sheets (uses <tbody><td> everywhere but bolds
        //   header cells when the user has formatted them as headers)
        var hasHeader = table.querySelector('thead') !== null ||
            (rows[0] && Array.from(rows[0].children).some(function (child) {
                return child.tagName && child.tagName.toLowerCase() === 'th';
            })) ||
            (rows[0] && rowIsBold(rows[0]));

        var grid = rows.map(function (row) {
            return Array.from(row.children).filter(function (child) {
                var tag = child.tagName && child.tagName.toLowerCase();
                return tag === 'td' || tag === 'th';
            }).reduce(function (acc, cell) {
                var span = parseInt(cell.getAttribute('colspan'), 10) || 1;
                var text = cellText(cell);
                for (var i = 0; i < span; i++) acc.push(text);
                return acc;
            }, []);
        });

        // Normalise all rows to the same column count
        var cols = Math.max.apply(null, grid.map(function (r) { return r.length; }));
        if (!cols) return '';

        grid = grid.map(function (row) {
            while (row.length < cols) row.push('');
            return row;
        });

        var sep = '|' + Array(cols).fill(' --- |').join('');

        if (hasHeader) {
            // First row is a real header — promote it to the GFM header position.
            var header = '| ' + grid[0].join(' | ') + ' |';
            var body = grid.slice(1).map(function (row) {
                return '| ' + row.join(' | ') + ' |';
            }).join('\n');
            return body ? header + '\n' + sep + '\n' + body
                : header + '\n' + sep;
        } else {
            // No semantic header — emit blank header cells so no data row is
            // wrongly promoted (GFM tables require a header row and separator).
            var blankHeader = '|' + Array(cols).fill('  |').join('');
            var body = grid.map(function (row) {
                return '| ' + row.join(' | ') + ' |';
            }).join('\n');
            return body ? blankHeader + '\n' + sep + '\n' + body
                : blankHeader + '\n' + sep;
        }
    }

    function convertHtmlWithTables(html) {
        var cleaned = cleanHtml(html);
        var doc = new DOMParser().parseFromString(cleaned, 'text/html');
        var tables = Array.from(doc.querySelectorAll('table'));
        if (!tables.length) return getTurndown().turndown(cleaned);

        var markerBase = 'C2MTABLETOKEN' + Date.now().toString(36) + Math.random().toString(36).slice(2).toUpperCase();
        var replacements = [];

        tables.forEach(function (table, index) {
            if (!table.parentNode) return;

            var token = markerBase + '_' + index + '_';
            var markdown = tableToMarkdown(table);
            var markerNode = doc.createTextNode('\n' + token + '\n');
            table.parentNode.replaceChild(markerNode, table);
            replacements.push({ token: token, markdown: markdown });
        });

        if (!replacements.length) return getTurndown().turndown(cleaned);

        var markdown = getTurndown().turndown(doc.body.innerHTML);
        var matchedTokens = replacements.filter(function (item) {
            return markdown.indexOf(item.token) !== -1;
        }).length;

        if (matchedTokens !== replacements.length) {
            return getTurndown().turndown(cleaned);
        }

        replacements.forEach(function (item) {
            var replacement = item.markdown ? ('\n\n' + item.markdown + '\n\n') : '\n\n';
            markdown = markdown.split(item.token).join(replacement);
        });

        return markdown.replace(/\n{3,}/g, '\n\n').trim();
    }

    // ---------------------------------------------------------------------------
    // General HTML → markdown (via Turndown) with light pre-processing
    // ---------------------------------------------------------------------------

    // Strip <style>/<script> blocks before handing HTML to Turndown so stray
    // CSS text doesn't end up in the output.
    function cleanHtml(html) {
        var doc = new DOMParser().parseFromString(html, 'text/html');
        doc.querySelectorAll('style, script, link, meta').forEach(function (el) {
            el.remove();
        });
        return doc.body.innerHTML;
    }

    // ---------------------------------------------------------------------------
    // Main conversion entry point
    // ---------------------------------------------------------------------------

    function convertHtmlToMarkdown(html) {
        return convertHtmlWithTables(html);
    }

    // ---------------------------------------------------------------------------
    // UI
    // ---------------------------------------------------------------------------

    function showOutput(markdown) {
        var wrapper = document.getElementById('wrapper');
        var output = document.getElementById('output');
        wrapper.classList.remove('hidden');
        output.value = markdown;
    }

    function handlePaste(e) {
        var clipboardData = e.clipboardData || window.clipboardData;
        var html = clipboardData.getData('text/html');
        var text = clipboardData.getData('text/plain');

        var markdown = html ? convertHtmlToMarkdown(html) : text;

        var pastebin = document.getElementById('pastebin');
        setTimeout(function () { pastebin.innerHTML = ''; }, 0);

        showOutput(markdown);
        e.preventDefault();
    }

    function setupCopyButton() {
        var btn = document.getElementById('copy-btn');
        var status = document.getElementById('status');

        btn.addEventListener('click', function () {
            var output = document.getElementById('output');
            output.select();
            navigator.clipboard.writeText(output.value).then(function () {
                status.textContent = 'Copied!';
                setTimeout(function () { status.textContent = ''; }, 2000);
            }).catch(function () {
                document.execCommand('copy');
                status.textContent = 'Copied!';
                setTimeout(function () { status.textContent = ''; }, 2000);
            });
        });
    }

    document.addEventListener('DOMContentLoaded', function () {
        var pastebin = document.getElementById('pastebin');
        pastebin.addEventListener('paste', handlePaste);

        document.addEventListener('paste', function (e) {
            if (e.target !== pastebin) handlePaste(e);
        });

        pastebin.focus();
        setupCopyButton();
    });
})();
