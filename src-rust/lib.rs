//! Papyr DOCX Parser Library
//! 
//! A lightweight, read-only DOCX parser focused on extracting structured content
//! for rendering in a Tauri-based document viewer.

pub mod models;
pub mod comment_parser;

pub use models::*;
pub use comment_parser::CommentParser;

use std::io::BufRead;

/// Main entry point for parsing a DOCX file
pub struct DocxParser {
    comment_parser: CommentParser,
}

impl DocxParser {
    pub fn new() -> Self {
        Self {
            comment_parser: CommentParser::new(),
        }
    }

    /// Parse comments from a DOCX comments.xml file
    pub fn parse_comments<R: BufRead>(&mut self, reader: R) -> Result<Vec<Comment>, Box<dyn std::error::Error>> {
        self.comment_parser.parse_comments_xml(reader)?;
        Ok(self.comment_parser.get_comments())
    }

    /// Get a reference to the comment parser for document parsing integration
    pub fn comment_parser(&mut self) -> &mut CommentParser {
        &mut self.comment_parser
    }
}

impl Default for DocxParser {
    fn default() -> Self {
        Self::new()
    }
}