<?php
/**
 * Plugin Name: Markdown to DOCX Converter
 * Description: Convert Markdown text to Word DOCX format
 * Version: 1.0
 * Author: Damon Noisette
 */
// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

// Load PHPWord if available
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}

class MarkdownToDocxConverter {
    public function __construct() {
        add_action('admin_menu', array($this, 'add_admin_menu'));
        add_action('admin_post_convert_markdown', array($this, 'handle_conversion'));
    }

    public function add_admin_menu() {
        add_submenu_page(
            'tools.php',                  // parent slug (Tools)
            'Markdown to DOCX',           // page title
            'MD to DOCX',                 // menu title
            'manage_options',             // capability
            'markdown-to-docx',           // menu slug
            array($this, 'admin_page')    // callback function
        );
    }

    public function admin_page() {
        ?>
        <div class="wrap">
            <h1>Markdown to DOCX Converter</h1>
            <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
                <input type="hidden" name="action" value="convert_markdown">
                <?php wp_nonce_field('convert_markdown_nonce', 'markdown_nonce'); ?>
                <table class="form-table">
                    <tr>
                        <th scope="row">
                            <label for="markdown_content">Markdown Content</label>
                        </th>
                        <td>
                            <textarea 
                                name="markdown_content" 
                                id="markdown_content" 
                                rows="15" 
                                cols="80" 
                                class="large-text"
                                placeholder="Enter your Markdown content here..."
                            ></textarea>
                            <p class="description">
                                Enter your Markdown text above. Supports headers, bold, italic, links, lists, tables, and code blocks.
                            </p>
                        </td>
                    </tr>
                    <tr>
                        <th scope="row">
                            <label for="document_title">Document Title</label>
                        </th>
                        <td>
                            <input 
                                type="text" 
                                name="document_title" 
                                id="document_title" 
                                value="Converted Document" 
                                class="regular-text"
                            >
                            <p class="description">
                                This will be used as the filename for your DOCX file.
                            </p>
                        </td>
                    </tr>
                </table>
                <?php submit_button('Convert to DOCX', 'primary', 'submit', false); ?>
            </form>
            <div class="markdown-help">
                <h3>Supported Markdown Syntax</h3>
                <p><strong>Headers:</strong> # H1, ## H2, ### H3</p>
                <p><strong>Bold:</strong> **bold text**</p>
                <p><strong>Italic:</strong> *italic text*</p>
                <p><strong>Links:</strong> [link text](URL)</p>
                <p><strong>Lists:</strong> - item or 1. item</p>
                <p><strong>Code:</strong> `inline code` or ```code block```</p>
                <p><strong>Tables:</strong> Use pipes and dashes like below</p>
<pre>
| Header 1 | Header 2 |
|----------|----------|
| Row A1   | Row A2   |
| Row B1   | Row B2   |
</pre>
            </div>
        </div>
        <style>
        .markdown-help {
            margin-top: 30px;
            padding: 20px;
            background: #f9f9f9;
            border-left: 4px solid #0073aa;
        }
        .markdown-help h3 {
            margin-top: 0;
        }
        .markdown-help p {
            margin: 5px 0;
        }
        pre {
            background: #eee;
            padding: 10px;
            overflow-x: auto;
        }
        </style>
        <?php
    }

    public function handle_conversion() {
        // Verify nonce
        if (!isset($_POST['markdown_nonce']) || !wp_verify_nonce($_POST['markdown_nonce'], 'convert_markdown_nonce')) {
            wp_die('Security check failed');
        }

        // Check user permissions
        if (!current_user_can('manage_options')) {
            wp_die('Insufficient permissions');
        }

        $markdown_content = sanitize_textarea_field($_POST['markdown_content']);
        $document_title = sanitize_text_field($_POST['document_title']);

        if (empty($markdown_content)) {
            wp_redirect(admin_url('admin.php?page=markdown-to-docx&error=empty'));
            exit;
        }

        // Convert Markdown to HTML
        $html_content = $this->markdown_to_html($markdown_content);

        // Create DOCX file
        $this->create_docx_file($html_content, $document_title);
    }

    private function markdown_to_html($markdown) {
        $html = $markdown;

        // Normalize line breaks
        $html = preg_replace('/\r\n|\r/', "\n", $html);

        // Convert tables
        $html = $this->convert_markdown_tables($html);

        // Headers
        $html = preg_replace('/^### (.*$)/m', '<h3>$1</h3>', $html);
        $html = preg_replace('/^## (.*$)/m', '<h2>$1</h2>', $html);
        $html = preg_replace('/^# (.*$)/m', '<h1>$1</h1>', $html);

        // Bold and Italic
        $html = preg_replace('/\*\*(.*?)\*\*/', '<strong>$1</strong>', $html);
        $html = preg_replace('/\*(.*?)\*/', '<em>$1</em>', $html);

        // Links
        $html = preg_replace('/$$([^$$]+)$$$([^$$]+)/', '<a href="$2">$1</a>', $html);

        // Code blocks
        $html = preg_replace('/```(.*?)```/s', '<pre><code>$1</code></pre>', $html);
        $html = preg_replace('/`([^`]+)`/', '<code>$1</code>', $html);

        // Unordered lists
        $html = preg_replace_callback('/(?:^- .+\n?)+/m', function ($matches) {
            $items = preg_replace('/^- (.+)$/m', '<li>$1</li>', $matches[0]);
            return "<ul>{$items}</ul>";
        }, $html);

        // Ordered lists
        $html = preg_replace_callback('/(?:^\d+\. .+\n?)+/m', function ($matches) {
            $items = preg_replace('/^(\d+)\. (.+)$/m', '<li>$2</li>', $matches[0]);
            return "<ol>{$items}</ol>";
        }, $html);

        // Paragraphs
        $html = preg_replace('/\n{2,}/', '</p><p>', $html);
        $html = '<p>' . trim($html) . '</p>';

        // Clean up empty paragraphs
        $html = preg_replace('/<p>\s*<\/p>/', '', $html);

        return $html;
    }

    private function convert_markdown_tables($text) {
        return preg_replace_callback(
            '/^(\|.*\|)\n(\|.*\|)(?:\n((?:\|?.*\|?)+))?/m',
            function ($matches) {
                $headerRow = $this->parseTableRow($matches[1]);
                $separatorRow = $matches[2]; // Not used but ensures table format
                $bodyRows = isset($matches[3]) ? explode("\n", $matches[3]) : [];

                $html = '<table border="1" cellpadding="5" cellspacing="0"><thead><tr>';
                foreach ($headerRow as $cell) {
                    $html .= "<th>" . htmlspecialchars(trim($cell)) . "</th>";
                }
                $html .= '</tr></thead><tbody>';

                foreach ($bodyRows as $row) {
                    if (trim($row) === '') continue;
                    $cells = $this->parseTableRow($row);
                    $html .= '<tr>';
                    foreach ($cells as $cell) {
                        $html .= "<td>" . htmlspecialchars(trim($cell)) . "</td>";
                    }
                    $html .= '</tr>';
                }

                $html .= '</tbody></table>';
                return $html;
            },
            $text
        );
    }

    private function parseTableRow($row) {
        $cells = array_map('trim', explode('|', trim($row)));
        array_pop($cells); // Remove last empty element
        array_shift($cells); // Remove first empty element
        return $cells;
    }

    private function create_docx_file($html_content, $title) {
        if (class_exists('\PhpOffice\PhpWord\PhpWord')) {
            $this->create_with_phpword($html_content, $title);
        } else {
            $this->create_word_file($html_content, $title);
        }
    }

    private function create_with_phpword($html_content, $title) {
        try {
            $phpWord = new \PhpOffice\PhpWord\PhpWord();

            // Set document properties
            $properties = $phpWord->getDocInfo();
            $properties->setCreator('WordPress Markdown to DOCX Plugin');
            $properties->setTitle($title);

            // Add content
            \PhpOffice\PhpWord\Shared\Html::addHtml($phpWord->addSection(), $html_content);

            // Save as DOCX
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Disposition: attachment;filename="' . sanitize_file_name($title) . '.docx"');
            header('Cache-Control: max-age=0');

            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $objWriter->save('php://output');
            exit;
        } catch (\Exception $e) {
            // Fallback to Word-compatible HTML doc
            $this->create_word_file($html_content, $title);
        }
    }

    private function create_word_file($html_content, $title) {
        header("Content-Type: application/msword");
        header("Content-Disposition: attachment;Filename=" . sanitize_file_name($title) . ".doc");
        header("Pragma: no-cache");
        header("Expires: 0");

        echo "<html><head><meta charset='UTF-8'><title>" . esc_html($title) . "</title></head><body>";
        echo $html_content;
        echo "</body></html>";

        exit;
    }
}

// Initialize the plugin
function run_markdown_to_docx_converter() {
    new MarkdownToDocxConverter();
}
add_action('plugins_loaded', 'run_markdown_to_docx_converter');