# 📝 University Paper Writer

**AI-Powered Academic Writing Assistant with Claude Integration**

A sophisticated desktop application that harnesses the power of Anthropic's Claude API to generate high-quality university papers, essays, and academic documents. Built with Python and Tkinter, this tool streamlines the academic writing process while maintaining proper formatting, citations, and scholarly standards.

## ✨ Key Features

### 🤖 **Claude API Integration**
- **Advanced AI Writing**: Leverages Claude's sophisticated language model for academic content generation
- **Intelligent Paper Structure**: Automatically organizes content into proper academic format
- **Context-Aware Generation**: Incorporates uploaded scripts and materials into coherent arguments
- **Customizable Instructions**: Define specific requirements, word counts, and academic styles

### 📚 **Comprehensive Script Management**
- **Multi-Format Support**: Upload and process PDF, DOCX, and text files
- **Script Organization**: Manage multiple source documents with priority ordering
- **Content Integration**: Seamlessly weave uploaded materials into generated papers
- **Reference Extraction**: Automatically extract and utilize key information from sources

### 📝 **Professional Document Formatting**
- **University-Standard Layout**: Proper margins, fonts, and spacing for academic submissions
- **Automated Header Generation**: Creates professional title pages with student information
- **Structured Content**: Introduction, body paragraphs, and conclusion with proper flow
- **Citation Integration**: Harvard-style in-text citations and reference formatting
- **Word Document Export**: Save papers in universally compatible .docx format

### ⚙️ **Advanced Configuration**
- **Custom Instructions**: Define paper requirements, topics, and specific guidelines
- **API Management**: Secure storage and configuration of Claude API credentials
- **Settings Persistence**: All configurations saved automatically across sessions
- **Error Handling**: Robust error recovery and user feedback systems

## 🚀 Quick Start Guide

### Prerequisites

1. **Python 3.8+** installed on your system
2. **Claude API Key** from Anthropic
3. **Required Libraries**: All dependencies listed in requirements

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/gambadio/University-Paper-Writer.git
   cd University-Paper-Writer
   ```

2. **Install required dependencies**:
   ```bash
   pip install requests python-docx PyPDF2 beautifulsoup4 markdown
   ```

3. **Launch the application**:
   ```bash
   python claude_paper_generator.py
   ```

4. **Configure API Settings**:
   - Enter your Claude API key in the settings panel
   - Configure your personal details (name, date, etc.)
   - Set default paper formatting preferences

## 📖 Comprehensive Usage Guide

### Setting Up Your First Paper

1. **Configure API Access**:
   - Launch the application and access the settings
   - Enter your valid Claude API key from Anthropic
   - Configure personal information for paper headers

2. **Upload Source Materials**:
   - Use "Manage Scripts" to upload relevant documents
   - Support for PDF lecture notes, research papers, and text files
   - Organize scripts by priority for optimal content integration

3. **Define Paper Instructions**:
   - Specify the paper topic, requirements, and word count
   - Include specific guidelines, citation styles, and formatting needs
   - Add any special instructions or academic standards to follow

4. **Generate Your Paper**:
   - Click "Generate Paper with Claude" to begin the AI writing process
   - Monitor progress as Claude processes your materials and instructions
   - Review the generated content for accuracy and completeness

5. **Export and Finalize**:
   - Save your completed paper as a formatted Word document
   - Review formatting, citations, and overall structure
   - Make any necessary manual adjustments before submission

### Advanced Features

**Script Management System**:
- Upload multiple source documents and reference materials
- Reorder scripts by priority to influence content integration
- Delete or modify uploaded materials as needed
- Automatic text extraction from various file formats

**Instruction Customization**:
- Create detailed paper requirements and academic guidelines
- Specify citation styles (Harvard, APA, MLA)
- Define word count ranges and structural requirements
- Include assessment criteria and marking rubrics

**Professional Formatting**:
- Automatic title page generation with student details
- Proper academic heading hierarchy and numbering
- Standard university margins, fonts, and line spacing
- Professional reference list formatting

## 🏗️ Technical Architecture

### Core Components

**Main Application (`claude_paper_generator.py`)**
- Tkinter-based GUI with professional academic styling
- Multi-window system for different management functions
- Integrated Claude API communication and response handling
- Comprehensive error handling and user feedback

**Document Processing Engine**
- Multi-format file reading (PDF, DOCX, TXT)
- Intelligent text extraction and content preparation
- Academic formatting with proper document structure
- Citation integration and reference management

**API Integration System**
- Secure Claude API communication
- Intelligent prompt construction for academic writing
- Response processing and content formatting
- Rate limiting and error recovery mechanisms

### Dependencies

- **requests**: HTTP communication with Claude API
- **python-docx**: Professional Word document generation and formatting
- **PyPDF2**: PDF text extraction and processing
- **beautifulsoup4**: HTML content parsing and cleanup
- **markdown**: Markdown processing for content formatting

## 📋 Paper Types & Academic Standards

### Supported Academic Formats

- **Research Essays**: Comprehensive argumentative papers with literature review
- **Case Study Analysis**: Detailed analytical papers with evidence-based conclusions
- **Literature Reviews**: Systematic analysis of academic sources and research
- **Argumentative Essays**: Structured persuasive writing with proper evidence support
- **Comparative Analysis**: Multi-perspective examination of topics or theories
- **Critical Reviews**: Analytical evaluation of academic works or concepts

### Academic Quality Features

- **Scholarly Tone**: Appropriate academic language and vocabulary level
- **Proper Structure**: Clear introduction, body, and conclusion organization
- **Citation Integration**: Seamless incorporation of references and quotes
- **Evidence-Based Arguments**: Logical reasoning supported by credible sources
- **Professional Formatting**: University-standard layout and presentation

## ⚙️ Configuration Options

### API Settings
- **Claude Model Selection**: Choose appropriate model for paper complexity
- **Token Management**: Optimize API usage for cost-effective generation
- **Response Parameters**: Fine-tune creativity and accuracy settings
- **Error Handling**: Configure retry attempts and fallback strategies

### Formatting Controls
- **Document Layout**: Margins, fonts, and spacing configuration
- **Header Styles**: Multi-level heading formatting and numbering
- **Citation Format**: Harvard, APA, or custom citation style selection
- **Export Options**: Word document compatibility and formatting preservation

### Content Management
- **Source Integration**: Control how uploaded materials influence paper content
- **Instruction Processing**: Define how requirements are interpreted and applied
- **Quality Control**: Set parameters for academic standards and expectations
- **Reference Handling**: Automatic citation generation and bibliography creation

## 🛠️ Troubleshooting

### Common Issues

**API Connection Problems**:
- Verify Claude API key is correctly entered and valid
- Check internet connection and Anthropic service status
- Ensure sufficient API credits are available in your account

**Document Processing Errors**:
- Verify uploaded files are not corrupted or password-protected
- Check file permissions and accessibility
- Ensure supported file formats (PDF, DOCX, TXT)

**Formatting Issues**:
- Review document export settings and Word compatibility
- Check citation style configuration and requirements
- Verify academic formatting standards are properly applied

**Content Quality Concerns**:
- Refine instruction clarity and specificity
- Provide more detailed requirements and guidelines
- Review uploaded source materials for relevance and quality

## 🎯 Academic Applications

### Student Use Cases
- **Essay Assignments**: Structured academic essays with proper research integration
- **Research Papers**: Comprehensive papers with literature review and analysis
- **Case Studies**: Detailed analytical papers with evidence-based conclusions
- **Literature Reviews**: Systematic examination of academic sources and research

### Professional Applications
- **Academic Research**: Preliminary drafts for scholarly publications
- **Grant Proposals**: Well-structured funding applications and research proposals
- **Conference Papers**: Academic presentations and conference submissions
- **Thesis Chapters**: Structured chapters for graduate and undergraduate theses

## 🤝 Contributing

University Paper Writer welcomes contributions to enhance academic writing capabilities:

- **Feature Development**: Additional citation styles, formatting options, and export formats
- **UI/UX Improvements**: Enhanced user interface and workflow optimization
- **API Integration**: Support for additional AI providers and models
- **Academic Standards**: Extended support for different university requirements
- **Documentation**: User guides, tutorials, and academic writing best practices

## 📄 License

This project is licensed for **non-commercial use only**.

**Personal Use**: Free to use, copy, and modify for private, educational, and non-commercial purposes.  
**Commercial Use**: Prohibited without explicit written permission.

For commercial licensing inquiries, please contact: **ricardo.kupper@adalala.com**

## ⚠️ Academic Integrity Notice

This tool is designed to assist with academic writing and research. Users are responsible for:
- Following their institution's academic integrity policies
- Properly citing all sources and AI assistance
- Using generated content as a starting point for original work
- Ensuring compliance with assignment requirements and academic standards

## 🌟 Acknowledgments

- **Anthropic**: For providing the Claude API that powers intelligent academic writing
- **Academic Community**: For inspiration and requirements that shaped this educational tool
- **Python Libraries**: For the robust foundation that enables document processing and API integration
- **Educational Institutions**: For the academic standards and formats that guide this application

---

*Empowering academic excellence through intelligent writing assistance - where AI meets scholarly research.*

**© 2025 Ricardo Kupper** • Designed for students and researchers pursuing academic excellence
