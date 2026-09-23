use papyr::DocxParser;
use std::io::Cursor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Papyr DOCX Comment Parser Demo");
    
    // Example comments.xml content for demonstration
    let comments_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="John Doe" w:date="2023-01-15T10:30:00Z">
    <w:p>
      <w:r>
        <w:t>This is an important note about the document</w:t>
      </w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Jane Smith" w:date="2023-01-16T14:45:00Z">
    <w:p>
      <w:r>
        <w:t>Please review this section carefully</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>"#;

    let mut parser = DocxParser::new();
    let cursor = Cursor::new(comments_xml);
    
    match parser.parse_comments(cursor) {
        Ok(comments) => {
            println!("\nParsed {} comments:", comments.len());
            for comment in comments {
                println!("  Comment {}: {} ({})", 
                    comment.id, 
                    comment.author, 
                    comment.date.unwrap_or_else(|| "No date".to_string())
                );
                println!("    Text: {}", comment.text);
                println!();
            }
        }
        Err(e) => {
            eprintln!("Error parsing comments: {}", e);
        }
    }

    Ok(())
}