<?php
/**
 * Plugin Name: Markdown to DOCX Converter
 * Description: Convert Markdown text to Word DOCX format. Use [markdown_to_docx] shortcode to embed on any page.
 * Version: 1.6
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

/**
 * Class MarkdownToDocxConverter
 *
 * Main plugin class responsible for:
 * - Admin interface
 * - Shortcode rendering
 * - Markdown parsing
 * - Document generation
 */
class MarkdownToDocxConverter {
    /**
     * Constructor
     * Initialize hooks and register shortcode
     */
    public function __construct() {
        add_action('init', array($this, 'init'));
        add_shortcode('markdown_to_docx', array($this, 'render_shortcode'));
    }

    /**
     * WordPress init hook handler
     * Used to conditionally initialize admin features only when needed
     */
    public function init() {
        if (is_admin()) {
            add_action('admin_init', array($this, 'admin_init'));
        }
    }

    /**
     * WordPress admin_init hook handler
     * Registers admin menu and form submission handler
     */
    public function admin_init() {
        add_action('admin_menu', array($this, 'add_admin_menu'));
        add_action('admin_post_convert_markdown', array($this, 'handle_conversion'));
    }

    /**
     * Adds the admin menu page under Tools or top-level menu
     */
    public function add_admin_menu() {
        add_menu_page(
            'Markdown to DOCX',
            'MD to DOCX',
            'manage_options',
            'markdown-to-docx',
            array($this, 'admin_page'),
            'dashicons-edit-page',
            30
        );
    }

    /**
     * Renders the admin settings page UI
     */
    public function admin_page() {
        $this->render_converter_form();
    }

    /**
     * Renders the Markdown conversion form via shortcode
     * @return string HTML output of the form
     */
    public function render_shortcode() {
        ob_start(); // Start buffering output
        $this->render_converter_form(); // Render the form
        return ob_get_clean(); // Return buffer content as string
    }

    /**
     * Renders the shared converter form (used in both admin and frontend)
     */
    private function render_converter_form() {
        ?>
        <div class="markdown-to-docx-converter-container">
            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="markdown-to-docx-form">
                <input type="hidden" name="action" value="convert_markdown">
                <?php wp_nonce_field('convert_markdown_nonce', 'markdown_nonce'); ?>
                <label for="markdown_content">Enter Markdown:</label>
                <textarea name="markdown_content" id="markdown_content" rows="10" cols="80"
                          placeholder="Type your Markdown here..."><?php
                    echo isset($_POST['markdown_content']) ? esc_textarea(stripslashes($_POST['markdown_content'])) : '';
                    ?></textarea>
                <label for="document_title">Document Title:</label>
                <input type="text" name="document_title" id="document_title"
                       value="<?php echo isset($_POST['document_title']) ? esc_attr($_POST['document_title']) : 'Converted_Document'; ?>">
                <input type="submit" name="submit" class="button button-primary" value="Convert to DOCX">
            </form>
            <div class="markdown-syntax-help">
                <h4>Supported Syntax</h4>
                <ul>
                    <li><strong>Headers:</strong> # H1, ## H2, ### H3, #### H4</li>
                    <li><strong>Bold:</strong> **bold text**</li>
                    <li><strong>Italic:</strong> *italic text*</li>
                    <li><strong>Links:</strong> [text](url)</li>
                    <li><strong>Lists:</strong> - item or 1. item</li>
                    <li><strong>Code:</strong> `inline` or ```block```</li>
                    <li><strong>Tables:</strong> | Column1 | Column2 |</li>
                </ul>
            </div>
        </div>
        <!-- Isolated CSS to prevent theme conflicts -->
        <style>
            .markdown-to-docx-converter-container {
                max-width: 800px;
                margin: 2rem auto;
                padding: 1.5rem;
                background-color: #fff;
                border: 1px solid #ddd;
                border-radius: 8px;
                font-family: sans-serif;
            }
            .markdown-to-docx-form label {
                display: block;
                margin-top: 1rem;
                font-weight: bold;
            }
            .markdown-to-docx-form textarea,
            .markdown-to-docx-form input[type="text"] {
                width: 100%;
                padding: 0.5rem;
                margin-top: 0.3rem;
                box-sizing: border-box;
                border: 1px solid #ccc;
                border-radius: 4px;
                font-size: 1rem;
                font-family: monospace;
            }
            .markdown-to-docx-form input[type="submit"] {
                margin-top: 1rem;
                padding: 0.5rem 1rem;
            }
            .markdown-syntax-help {
                margin-top: 1.5rem;
                padding: 1rem;
                background: #f9f9f9;
                border-left: 4px solid #0073aa;
                border-radius: 4px;
            }
            .markdown-syntax-help ul {
                margin: 0;
                padding-left: 1.2rem;
            }
            .markdown-syntax-help li {
                margin-bottom: 0.4rem;
            }
        </style>
        <?php
    }

    /**
     * Handles form submission after user clicks "Convert to DOCX"
     */
    public function handle_conversion() {
        // Verify security nonce
        if (!isset($_POST['markdown_nonce']) || !wp_verify_nonce($_POST['markdown_nonce'], 'convert_markdown_nonce')) {
            wp_die('Security check failed');
        }
        // Restrict to users who can edit posts
        if (!current_user_can('edit_posts')) {
            wp_die('You are not allowed to convert documents.');
        }
        // Sanitize and unescape markdown content
        $markdown_content = stripslashes(sanitize_textarea_field($_POST['markdown_content']));
        $document_title = stripslashes(sanitize_text_field($_POST['document_title']));
        if (empty($markdown_content)) {
            wp_redirect(admin_url('admin.php?page=markdown-to-docx&error=empty'));
            exit;
        }
        // Convert Markdown to HTML
        $html_content = $this->markdown_to_html($markdown_content);
        // Generate and download the final document
        $this->create_docx_file($html_content, $document_title);
    }

    /**
     * Converts basic Markdown syntax into HTML
     * Supports headers, bold, italic, links, lists, paragraphs, and tables
     * @param string $markdown Raw Markdown input
     * @return string HTML output
     */
    private function markdown_to_html($markdown) {
        $html = $markdown;

        // Normalize line breaks across platforms
        $html = preg_replace('/\r\n|\r/', "\n", $html);

        // Headers
        $html = preg_replace('/^#### (.*)$/m', '<h4>$1</h4>', $html); // H4 support
        $html = preg_replace('/^### (.*)$/m', '<h3>$1</h3>', $html);
        $html = preg_replace('/^## (.*)$/m', '<h2>$1</h2>', $html);
        $html = preg_replace('/^# (.*)$/m', '<h1>$1</h1>', $html);

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

        // Tables
        $html = preg_replace_callback('/^(?:\|.*?\|\n){2,}/m', function ($matches) {
            $lines = explode("\n", trim($matches[0]));
            $header = str_getcsv(trim($lines[0], '|'), '|');
            $separator = str_getcsv(trim($lines[1], '|'), '|');

            // Validate separator line
            foreach ($separator as $cell) {
                if (!preg_match('/^-+$/', $cell)) {
                    return $matches[0]; // Invalid table structure
                }
            }

            $rows = array_map(function ($line) {
                return str_getcsv(trim($line, '|'), '|');
            }, array_slice($lines, 2));

            // Build HTML table
            $table = '<table>';
            $table .= '<thead><tr>';
            foreach ($header as $cell) {
                $table .= '<th>' . htmlspecialchars($cell) . '</th>';
            }
            $table .= '</tr></thead>';
            $table .= '<tbody>';

            foreach ($rows as $row) {
                $table .= '<tr>';
                foreach ($row as $cell) {
                    // Apply bold/italic formatting within cells
                    $cell = preg_replace('/\*\*(.*?)\*\*/', '<strong>$1</strong>', $cell);
                    $cell = preg_replace('/\*(.*?)\*/', '<em>$1</em>', $cell);
                    $table .= '<td>' . htmlspecialchars($cell) . '</td>';
                }
                $table .= '</tr>';
            }

            $table .= '</tbody></table>';
            return $table;
        }, $html);

        // Paragraphs
        $html = preg_replace('/\n{2,}/', '</p><p>', $html);
        $html = '<p>' . trim($html) . '</p>';

        // Remove empty paragraphs
        $html = preg_replace('/<p>\s*<\/p>/', '', $html);

        return $html;
    }

    /**
     * Generates a Word-compatible document based on available libraries
     * Uses PHPWord if possible, falls back to simple .doc file
     * @param string $html_content HTML version of the Markdown content
     * @param string $title Document title used in filename
     */
    private function create_docx_file($html_content, $title) {
        if (class_exists('\PhpOffice\PhpWord\PhpWord')) {
            $this->create_with_phpword($html_content, $title);
        } else {
            $this->create_word_file($html_content, $title);
        }
    }

    /**
     * Uses PHPWord library to generate real DOCX file
     * @param string $html_content
     * @param string $title
     */
    private function create_with_phpword($html_content, $title) {
        try {
            $phpWord = new \PhpOffice\PhpWord\PhpWord();
            // Set document metadata
            $properties = $phpWord->getDocInfo();
            $properties->setCreator('WordPress Markdown to DOCX Plugin');
            $properties->setTitle($title);
            // Add content using PHPWord's built-in HTML parser
            \PhpOffice\PhpWord\Shared\Html::addHtml($phpWord->addSection(), $html_content);
            // Send headers to force download
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Disposition: attachment;filename="' . sanitize_file_name($title) . '.docx"');
            header('Cache-Control: max-age=0');
            // Save to output stream
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $objWriter->save('php://output');
            exit;
        } catch (\Exception $e) {
            // Fallback to Word-compatible HTML doc
            $this->create_word_file($html_content, $title);
        }
    }

    /**
     * Fallback method to generate simple .doc file using HTML
     * Works by wrapping HTML in Word-readable tags
     * @param string $html_content
     * @param string $title
     */
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

// Instantiate the plugin once WordPress has loaded all dependencies
add_action('plugins_loaded', function () {
    new MarkdownToDocxConverter();
});
?>
