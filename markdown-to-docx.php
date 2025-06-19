<?php
/**
 * Plugin Name: Markdown to DOCX Converter
 * Description: Convert Markdown text to Word DOCX format. Use [markdown_to_docx] shortcode to embed on any page.
 * Version: 1.1
 * Author: Damon Noisette
 */
// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

// Load PHPWord if available (requires Composer autoloader)
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}

class MarkdownToDocxConverter {
    public function __construct() {
        add_action('init', array($this, 'init'));
        add_shortcode('markdown_to_docx', array($this, 'render_shortcode'));
    }

    public function init() {
        if (is_admin()) {
            add_action('admin_init', array($this, 'admin_init'));
        }
    }

    public function admin_init() {
        add_action('admin_menu', array($this, 'add_admin_menu'));
        add_action('admin_post_convert_markdown', array($this, 'handle_conversion'));
    }

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

    public function admin_page() {
        $this->render_converter_form();
    }

    public function render_shortcode() {
        ob_start();
        $this->render_converter_form();
        return ob_get_clean();
    }

    private function render_converter_form() {
        ?>
        <div class="markdown-to-docx-converter">
            <h2>Convert Markdown to Word</h2>
            <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
                <input type="hidden" name="action" value="convert_markdown">
                <?php wp_nonce_field('convert_markdown_nonce', 'markdown_nonce'); ?>
                <p>
                    <label for="markdown_content">Markdown Content:</label><br>
                    <textarea 
                        name="markdown_content" 
                        id="markdown_content" 
                        rows="10" 
                        cols="80" 
                        style="width: 100%;"
                        placeholder="Enter your Markdown content here..."><?php echo isset($_POST['markdown_content']) ? esc_textarea(stripslashes($_POST['markdown_content'])) : ''; ?></textarea>
                </p>
                <p>
                    <label for="document_title">Document Title:</label><br>
                    <input 
                        type="text" 
                        name="document_title" 
                        id="document_title" 
                        value="<?php echo isset($_POST['document_title']) ? esc_attr($_POST['document_title']) : 'Converted Document'; ?>" 
                        style="width: 100%;">
                </p>
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
            </div>
        </div>
        <style>
        .markdown-to-docx-converter {
            max-width: 800px;
            margin: 20px auto;
            padding: 20px;
            background: #fff;
            border: 1px solid #ddd;
            border-radius: 5px;
        }
        .markdown-help {
            margin-top: 20px;
            padding: 15px;
            background: #f9f9f9;
            border-left: 4px solid #0073aa;
        }
        .markdown-help h3 {
            margin-top: 0;
        }
        .markdown-help p {
            margin: 5px 0;
        }
        </style>
        <?php
    }

    public function handle_conversion() {
        // Verify nonce
        if (!isset($_POST['markdown_nonce']) || !wp_verify_nonce($_POST['markdown_nonce'], 'convert_markdown_nonce')) {
            wp_die('Security check failed');
        }

        // Optional: Restrict access
        if (!current_user_can('edit_posts')) {
            wp_die('You are not allowed to convert documents.');
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

// Initialize the plugin after WordPress is fully loaded
add_action('plugins_loaded', function () {
    new MarkdownToDocxConverter();
});
?>
