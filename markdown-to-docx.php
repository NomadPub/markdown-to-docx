<?php
/**
 * Plugin Name: Markdown to DOCX Converter
 * Description: Convert Markdown text to Word DOCX format. Use [markdown_to_docx] shortcode to embed on any page.
 * Version: 1.3
 * Author: Your Name
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
        // Initialize plugin
        add_action('plugins_loaded', array($this, 'init'));
    }

    /**
     * Initialize core functionality
     */
    public function init() {
        // Register shortcode
        add_shortcode('markdown_to_docx', array($this, 'render_shortcode'));

        // Admin-only hooks
        if (is_admin()) {
            add_action('admin_menu', array($this, 'add_submenu_page'));
            add_action('admin_post_convert_markdown', array($this, 'handle_conversion'));
        }
    }

    /**
     * Add submenu under Tools
     */
    public function add_submenu_page() {
        add_submenu_page(
            'tools.php',                     // Parent slug (Tools)
            'Markdown to DOCX Converter',    // Page title
            'Markdown to DOCX',              // Menu title
            'manage_options',                // Required capability
            'markdown-to-docx',              // Menu slug
            array($this, 'render_converter_form') // Callback
        );
    }

    /**
     * Shortcode handler
     */
    public function render_shortcode() {
        ob_start();
        $this->render_converter_form();
        return ob_get_clean();
    }

    /**
     * Render the HTML form
     */
    public function render_converter_form() {
        ?>
        <div class="markdown-to-docx-converter-container">
            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="markdown-to-docx-form">
                <input type="hidden" name="action" value="convert_markdown">
                <?php wp_nonce_field('convert_markdown_nonce', 'markdown_nonce'); ?>
                <label for="markdown_content">Enter Markdown:</label>
                <textarea name="markdown_content" id="markdown_content" rows="10" cols="80" placeholder="Type your Markdown here..."><?php
                    echo isset($_POST['markdown_content']) ? esc_textarea(stripslashes($_POST['markdown_content'])) : '';
                ?></textarea>
                <label for="document_title">Document Title:</label>
                <input type="text" name="document_title" id="document_title" value="<?php
                    echo isset($_POST['document_title']) ? esc_attr($_POST['document_title']) : 'Converted_Document';
                ?>">
                <input type="submit" name="submit" class="button button-primary" value="Convert to DOCX">
            </form>

            <!-- Syntax Help -->
            <div class="markdown-syntax-help">
                <h4>Supported Markdown Syntax</h4>
                <ul>
                    <li><strong>Headers:</strong> # H1, ## H2, ### H3</li>
                    <li><strong>Bold:</strong> **bold text**</li>
                    <li><strong>Italic:</strong> *italic text*</li>
                    <li><strong>Links:</strong> [text](url)</li>
                    <li><strong>Lists:</strong> - item or 1. item</li>
                    <li><strong>Code:</strong> `inline` or ```block```</li>
                </ul>
            </div>
        </div>

        <!-- Simple Styling -->
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
        </style>
        <?php
    }

    /**
     * Handle conversion form submission
     */
    public function handle_conversion() {
        // Security: Check nonce
        if (!isset($_POST['markdown_nonce']) || !wp_verify_nonce($_POST['markdown_nonce'], 'convert_markdown_nonce')) {
            wp_die(__('Security check failed.', 'markdown-to-docx'));
        }

        // Security: Check user permissions
        if (!current_user_can('manage_options')) {
            wp_die(__('You do not have permission to access this feature.', 'markdown-to-docx'));
        }

        // Sanitize input
        $markdown_content = sanitize_textarea_field($_POST['markdown_content']);
        $document_title   = sanitize_text_field($_POST['document_title']);

        // Validate content
        if (empty($markdown_content)) {
            wp_redirect(admin_url('admin.php?page=markdown-to-docx&error=empty'));
            exit;
        }

        // Convert Markdown to HTML
        $html_content = $this->markdown_to_html($markdown_content);

        // Generate and download file
        $this->create_docx_file($html_content, $document_title);
    }

    /**
     * Convert Markdown syntax to basic HTML
     */
    private function markdown_to_html($markdown) {
        $html = trim($markdown);
        $html = preg_replace('/\r\n|\r/', "\n", $html); // Normalize line breaks

        // Headers
        $html = preg_replace('/^### (.*$)/m', '<h3>$1</h3>', $html);
        $html = preg_replace('/^## (.*$)/m', '<h2>$1</h2>', $html);
        $html = preg_replace('/^# (.*$)/m', '<h1>$1</h1>', $html);

        // Bold / Italic
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
        $html = '<p>' . $html . '</p>';

        // Remove empty paragraphs
        $html = preg_replace('/<p>\s*<\/p>/', '', $html);

        return $html;
    }

    /**
     * Create DOCX or fallback to .doc
     */
    private function create_docx_file($html_content, $title) {
        if (class_exists('\PhpOffice\PhpWord\PhpWord')) {
            $this->create_with_phpword($html_content, $title);
        } else {
            $this->create_word_file($html_content, $title);
        }
    }

    /**
     * Create real DOCX using PHPWord
     */
    private function create_with_phpword($html_content, $title) {
        try {
            $phpWord = new \PhpOffice\PhpWord\PhpWord();

            // Set document info
            $properties = $phpWord->getDocInfo();
            $properties->setCreator('WordPress Markdown to DOCX Plugin');
            $properties->setTitle($title);

            // Add content
            \PhpOffice\PhpWord\Shared\Html::addHtml($phpWord->addSection(), $html_content);

            // Output headers
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Disposition: attachment;filename="' . sanitize_file_name($title) . '.docx"');
            header('Cache-Control: max-age=0');

            // Save and output
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $objWriter->save('php://output');
            exit;

        } catch (\Exception $e) {
            // Fallback to simple HTML doc
            $this->create_word_file($html_content, $title);
        }
    }

    /**
     * Fallback: Create simple HTML file that Word can open
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

// Start the plugin
new MarkdownToDocxConverter();
