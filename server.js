// server.js - Simple backend for AI Presentation Generator
const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const multer = require('multer');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const PizZip = require('pizzip');

const app = express();
const PORT = process.env.PORT || 3001;

// Configure multer for file uploads
const upload = multer({ 
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
      cb(null, true);
    } else {
      cb(new Error('Only PPTX files are allowed'), false);
    }
  },
  limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public')); // Serve frontend files

// AI Provider configurations
const providers = {
  openai: {
    url: 'https://api.openai.com/v1/chat/completions',
    headers: (apiKey) => ({
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    }),
    body: (prompt) => ({
      model: 'gpt-3.5-turbo', // Using the most reliable and available model
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.7,
      max_tokens: 4000
    }),
    extractResponse: (data) => {
      if (data.choices && data.choices[0] && data.choices[0].message) {
        return data.choices[0].message.content;
      }
      throw new Error('Invalid OpenAI response format');
    }
  },
  anthropic: {
    url: 'https://api.anthropic.com/v1/messages',
    headers: (apiKey) => ({
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    }),
    body: (prompt) => ({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4000,
      messages: [{ role: 'user', content: prompt }]
    }),
    extractResponse: (data) => {
      if (data.content && data.content[0] && data.content[0].text) {
        return data.content[0].text;
      }
      throw new Error('Invalid Anthropic response format');
    }
  },
  gemini: {
    url: (apiKey) => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`, // Using gemini-pro (stable and available)
    headers: () => ({
      'Content-Type': 'application/json'
    }),
    body: (prompt) => ({
      contents: [{ 
        parts: [{ text: prompt }] 
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 4000,
        topP: 0.8,
        topK: 10
      }
    }),
    extractResponse: (data) => {
      if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
        return data.candidates[0].content.parts[0].text;
      }
      throw new Error('Invalid Gemini response format');
    }
  }
};

// Main API endpoint for analyzing content
app.post('/api/analyze', async (req, res) => {
  try {
    const { text, guidance, provider, apiKey } = req.body;

    // Validation
    if (!text || !provider || !apiKey) {
      return res.status(400).json({ 
        error: 'Missing required fields: text, provider, and apiKey are required' 
      });
    }

    if (!providers[provider]) {
      return res.status(400).json({ 
        error: `Unsupported provider: ${provider}. Supported: ${Object.keys(providers).join(', ')}` 
      });
    }

    // Create the prompt
    const prompt = createAnalysisPrompt(text, guidance);
    console.log(`Making API call to ${provider}...`);

    // Get provider configuration
    const config = providers[provider];
    
    // Make API call
    const url = typeof config.url === 'function' ? config.url(apiKey) : config.url;
    
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: config.headers(apiKey),
        body: JSON.stringify(config.body(prompt))
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`${provider} API error:`, response.status, errorText);
        return res.status(response.status).json({ 
          error: `${provider} API error: ${response.status} ${response.statusText}`,
          details: errorText,
          provider: provider
        });
      }

      const data = await response.json();
      console.log(`${provider} API response received successfully`);
      
      const aiResponse = config.extractResponse(data);
      console.log(`${provider} response extracted, length:`, aiResponse?.length || 0);

      // Parse the AI response to extract JSON
      const slidesStructure = parseAIResponse(aiResponse);
      
      if (!slidesStructure) {
        return res.status(500).json({ 
          error: 'Failed to parse AI response into valid slide structure',
          rawResponse: aiResponse,
          provider: provider
        });
      }

      console.log(`${provider} analysis completed successfully, ${slidesStructure.slides?.length || 0} slides generated`);
      res.json({ success: true, slides: slidesStructure });

    } catch (fetchError) {
      console.error(`${provider} fetch error:`, fetchError.message);
      return res.status(500).json({ 
        error: `Network error calling ${provider} API`,
        details: fetchError.message,
        provider: provider
      });
    }

  } catch (error) {
    console.error('Analysis error:', error);
    res.status(500).json({ 
      error: 'Internal server error during analysis',
      details: error.message 
    });
  }
});

// New endpoint to generate and return PPTX file with optional template
app.post('/api/generate-pptx', upload.single('template'), async (req, res) => {
  try {
    const { text, guidance, provider, apiKey } = req.body;
    const templateFile = req.file; // This will contain the uploaded template file

    // Validation - template is now optional
    if (!text || !provider || !apiKey) {
      return res.status(400).json({ 
        error: 'Missing required fields: text, provider, and apiKey are required' 
      });
    }

    if (!providers[provider]) {
      return res.status(400).json({ 
        error: `Unsupported provider: ${provider}. Supported: ${Object.keys(providers).join(', ')}` 
      });
    }

    // Create the prompt
    const prompt = createAnalysisPrompt(text, guidance);

    // Get provider configuration
    const config = providers[provider];
    
    // Make API call
    const url = typeof config.url === 'function' ? config.url(apiKey) : config.url;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: config.headers(apiKey),
      body: JSON.stringify(config.body(prompt))
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`${provider} API error:`, response.status, errorText);
      return res.status(response.status).json({ 
        error: `${provider} API error: ${response.status} ${response.statusText}`,
        details: errorText
      });
    }

    const data = await response.json();
    const aiResponse = config.extractResponse(data);

    // Parse the AI response to extract JSON
    const slidesStructure = parseAIResponse(aiResponse);
    
    if (!slidesStructure) {
      return res.status(500).json({ 
        error: 'Failed to parse AI response into valid slide structure',
        rawResponse: aiResponse
      });
    }

    // Generate PPTX with optional template
    const pptxBuffer = await generatePPTX(slidesStructure, templateFile);
    
    // Set headers for file download
    const filename = `${slidesStructure.presentationTitle.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', pptxBuffer.length);
    
    // Send the PPTX file
    res.send(pptxBuffer);

  } catch (error) {
    console.error('PPTX generation error:', error);
    res.status(500).json({ 
      error: 'Internal server error during PPTX generation',
      details: error.message 
    });
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'healthy', timestamp: new Date().toISOString() });
});

// Function to read and extract information from PPTX template
async function readPPTXTemplate(templateBuffer) {
  try {
    const zip = new JSZip();
    const contents = await zip.loadAsync(templateBuffer);
    
    // Initialize template info object
    const templateInfo = {
      titleStyle: null,
      contentStyle: null,
      background: null,
      colorScheme: null,
      masterSlide: null
    };
    
    // Read theme colors from theme1.xml
    try {
      const themeFile = contents.files['ppt/theme/theme1.xml'];
      if (themeFile) {
        const themeXml = await themeFile.async('text');
        const colorScheme = extractColorScheme(themeXml);
        templateInfo.colorScheme = colorScheme;
        console.log('Extracted color scheme:', colorScheme);
      }
    } catch (error) {
      console.log('Could not read theme colors:', error.message);
    }
    
    // Read slide master for layout information
    try {
      const masterFile = contents.files['ppt/slideMasters/slideMaster1.xml'];
      if (masterFile) {
        const masterXml = await masterFile.async('text');
        const masterInfo = extractMasterSlideInfo(masterXml);
        templateInfo.masterSlide = masterInfo;
        console.log('Extracted master slide info');
      }
    } catch (error) {
      console.log('Could not read slide master:', error.message);
    }
    
    // Try to read the first slide for styling information
    try {
      const slideFile = contents.files['ppt/slides/slide1.xml'];
      if (slideFile) {
        const slideXml = await slideFile.async('text');
        const slideInfo = extractSlideInfo(slideXml, templateInfo.colorScheme);
        
        if (slideInfo.titleStyle) {
          templateInfo.titleStyle = slideInfo.titleStyle;
        }
        if (slideInfo.contentStyle) {
          templateInfo.contentStyle = slideInfo.contentStyle;
        }
        if (slideInfo.background) {
          templateInfo.background = slideInfo.background;
        }
        
        console.log('Extracted slide styling information');
      }
    } catch (error) {
      console.log('Could not read first slide:', error.message);
    }
    
    // Set default styles based on extracted information
    if (templateInfo.colorScheme && !templateInfo.titleStyle) {
      templateInfo.titleStyle = {
        fontSize: 28,
        fontFace: 'Arial',
        color: templateInfo.colorScheme.accent1 || '1F497D'
      };
    }
    
    if (templateInfo.colorScheme && !templateInfo.contentStyle) {
      templateInfo.contentStyle = {
        fontSize: 18,
        fontFace: 'Arial',
        color: templateInfo.colorScheme.text1 || '444444'
      };
    }
    
    return templateInfo;
    
  } catch (error) {
    console.error('Error reading PPTX template:', error);
    return null;
  }
}

// Helper function to extract color scheme from theme XML
function extractColorScheme(themeXml) {
  const colorScheme = {};
  
  try {
    // Simple regex patterns to extract common theme colors
    const patterns = {
      accent1: /<a:accent1[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:accent1>/i,
      accent2: /<a:accent2[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:accent2>/i,
      text1: /<a:dk1[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:dk1>/i,
      background1: /<a:lt1[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:lt1>/i
    };
    
    for (const [key, pattern] of Object.entries(patterns)) {
      const match = themeXml.match(pattern);
      if (match && match[1]) {
        colorScheme[key] = match[1];
      }
    }
    
    // Fallback colors if extraction fails
    if (!colorScheme.accent1) colorScheme.accent1 = '1F497D';
    if (!colorScheme.text1) colorScheme.text1 = '444444';
    if (!colorScheme.background1) colorScheme.background1 = 'FFFFFF';
    
  } catch (error) {
    console.log('Error extracting color scheme:', error.message);
    // Return default colors
    return {
      accent1: '1F497D',
      text1: '444444',
      background1: 'FFFFFF'
    };
  }
  
  return colorScheme;
}

// Helper function to extract master slide information
function extractMasterSlideInfo(masterXml) {
  // This is a simplified extraction - in a full implementation,
  // you would parse the XML properly to extract layout information
  return {
    extracted: true,
    hasCustomLayouts: masterXml.includes('<p:sldLayout')
  };
}

// Helper function to extract slide information
function extractSlideInfo(slideXml, colorScheme) {
  const slideInfo = {
    titleStyle: null,
    contentStyle: null,
    background: null
  };
  
  try {
    // Try to extract background information
    if (slideXml.includes('<p:bg>')) {
      const bgColorMatch = slideXml.match(/<a:srgbClr val="([^"]*)".*?>/);
      if (bgColorMatch) {
        slideInfo.background = {
          type: 'color',
          value: bgColorMatch[1]
        };
      }
    }
    
    // Use color scheme for styling if available
    if (colorScheme) {
      slideInfo.titleStyle = {
        fontSize: 28,
        fontFace: 'Arial',
        color: colorScheme.accent1
      };
      
      slideInfo.contentStyle = {
        fontSize: 18,
        fontFace: 'Arial',
        color: colorScheme.text1
      };
    }
    
  } catch (error) {
    console.log('Error extracting slide info:', error.message);
  }
  
  return slideInfo;
}

// Function to read template slides and extract their structure
async function readTemplateSlides(templateBuffer) {
  try {
    const zip = new JSZip();
    const contents = await zip.loadAsync(templateBuffer);
    
    const templateSlides = [];
    
    // Find all slide files
    const slideFiles = Object.keys(contents.files).filter(name => 
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );
    
    console.log(`Found ${slideFiles.length} slides in template`);
    
    for (const slideFileName of slideFiles) {
      try {
        const slideFile = contents.files[slideFileName];
        const slideXml = await slideFile.async('text');
        
        // Extract slide information
        const slideInfo = await extractSlideStructure(slideXml, contents);
        if (slideInfo) {
          templateSlides.push(slideInfo);
        }
      } catch (error) {
        console.log(`Error reading slide ${slideFileName}:`, error.message);
      }
    }
    
    return templateSlides.length > 0 ? templateSlides : null;
    
  } catch (error) {
    console.error('Error reading template slides:', error);
    return null;
  }
}

// Helper function to extract slide structure from XML
async function extractSlideStructure(slideXml, zipContents) {
  const slideInfo = {
    background: null,
    titlePlaceholder: null,
    contentPlaceholder: null,
    images: []
  };
  
  try {
    // Extract background information
    const backgroundMatch = slideXml.match(/<p:bg>(.*?)<\/p:bg>/s);
    if (backgroundMatch) {
      const bgContent = backgroundMatch[1];
      
      // Check for solid color background
      const colorMatch = bgContent.match(/<a:srgbClr val="([^"]*)".*?>/);
      if (colorMatch) {
        slideInfo.background = { color: colorMatch[1] };
      }
      
      // Check for image background
      const imageMatch = bgContent.match(/<a:blip r:embed="([^"]*)".*?>/);
      if (imageMatch) {
        slideInfo.background = { type: 'image', embed: imageMatch[1] };
      }
    }
    
    // Extract title placeholder position and formatting
    const titleMatches = slideXml.match(/<p:sp[^>]*>.*?<p:nvSpPr>.*?<p:cNvPr[^>]*name="Title[^"]*".*?<\/p:sp>/s);
    if (titleMatches) {
      slideInfo.titlePlaceholder = extractPlaceholderInfo(titleMatches[0]);
      slideInfo.titlePlaceholder.type = 'title';
    }
    
    // Extract content placeholder
    const contentMatches = slideXml.match(/<p:sp[^>]*>.*?<p:nvSpPr>.*?<p:cNvPr[^>]*name="Content Placeholder[^"]*".*?<\/p:sp>/s);
    if (contentMatches) {
      slideInfo.contentPlaceholder = extractPlaceholderInfo(contentMatches[0]);
      slideInfo.contentPlaceholder.type = 'content';
    }
    
    // If no specific content placeholder, look for text placeholders
    if (!slideInfo.contentPlaceholder) {
      const textMatches = slideXml.match(/<p:sp[^>]*>.*?<p:txBody>.*?<\/p:sp>/s);
      if (textMatches) {
        slideInfo.contentPlaceholder = extractPlaceholderInfo(textMatches[0]);
        slideInfo.contentPlaceholder.type = 'text';
      }
    }
    
    // Set default positions if not found
    if (!slideInfo.titlePlaceholder) {
      slideInfo.titlePlaceholder = {
        x: 0.5, y: 0.5, w: 9, h: 1,
        fontSize: 28, fontFace: 'Arial', color: '1F497D'
      };
    }
    
    if (!slideInfo.contentPlaceholder) {
      slideInfo.contentPlaceholder = {
        x: 0.5, y: 2, w: 9, h: 4,
        fontSize: 18, fontFace: 'Arial', color: '444444'
      };
    }
    
    return slideInfo;
    
  } catch (error) {
    console.log('Error extracting slide structure:', error.message);
    return null;
  }
}

// Helper function to extract placeholder information from XML
function extractPlaceholderInfo(xmlContent) {
  const placeholder = {
    x: 0.5, y: 0.5, w: 9, h: 1,
    fontSize: 18, fontFace: 'Arial', color: '444444'
  };
  
  try {
    // Extract position information (simplified)
    const transformMatch = xmlContent.match(/<a:xfrm>(.*?)<\/a:xfrm>/s);
    if (transformMatch) {
      const transform = transformMatch[1];
      
      // Extract offset (position)
      const offsetMatch = transform.match(/<a:off x="([^"]*)" y="([^"]*)"/);
      if (offsetMatch) {
        // Convert EMU to inches (approximate)
        placeholder.x = Math.round((parseInt(offsetMatch[1]) / 914400) * 100) / 100;
        placeholder.y = Math.round((parseInt(offsetMatch[2]) / 914400) * 100) / 100;
      }
      
      // Extract extent (size)
      const extentMatch = transform.match(/<a:ext cx="([^"]*)" cy="([^"]*)"/);
      if (extentMatch) {
        // Convert EMU to inches (approximate)
        placeholder.w = Math.round((parseInt(extentMatch[1]) / 914400) * 100) / 100;
        placeholder.h = Math.round((parseInt(extentMatch[2]) / 914400) * 100) / 100;
      }
    }
    
    // Extract font information
    const fontMatch = xmlContent.match(/<a:latin typeface="([^"]*)"/);
    if (fontMatch) {
      placeholder.fontFace = fontMatch[1];
    }
    
    // Extract font size
    const sizeMatch = xmlContent.match(/<a:rPr.*?sz="([^"]*)"/);
    if (sizeMatch) {
      placeholder.fontSize = Math.round(parseInt(sizeMatch[1]) / 100);
    }
    
    // Extract color
    const colorMatch = xmlContent.match(/<a:srgbClr val="([^"]*)"/);
    if (colorMatch) {
      placeholder.color = colorMatch[1];
    }
    
  } catch (error) {
    console.log('Error extracting placeholder info:', error.message);
  }
  
  return placeholder;
}

// Generate PPTX from slides structure with optional template support
// Generate PPTX from slides structure with template support
async function generatePPTX(slidesStructure, templateFile = null) {
  
  if (templateFile && templateFile.buffer) {
    // Use template-based approach - modify existing PPTX file
    console.log('Using template file for PPTX generation');
    return await generateFromTemplate(slidesStructure, templateFile.buffer);
  } else {
    // Use pptxgenjs for new presentations
    console.log('Creating new presentation without template');
    return await generateNewPresentation(slidesStructure);
  }
}

// Generate PPTX using existing template file
async function generateFromTemplate(slidesStructure, templateBuffer) {
  try {
    console.log('Using template to create NEW slides with template styling');
    
    // Option 1: Use pptxgenjs with template styling extracted
    return await generateNewSlidesWithTemplateStyle(slidesStructure, templateBuffer);
    
  } catch (error) {
    console.error('Error generating PPTX from template:', error);
    // Fallback to creating new presentation
    console.log('Falling back to new presentation generation');
    return await generateNewPresentation(slidesStructure);
  }
}

// Create new slides using template's styling but with our content
async function generateNewSlidesWithTemplateStyle(slidesStructure, templateBuffer) {
  try {
    // Extract theme and styling from template
    const templateTheme = await extractTemplateTheme(templateBuffer);
    console.log('Extracted template theme:', templateTheme ? 'Success' : 'Failed');
    
    // Create new presentation with extracted styling
    const pres = new PptxGenJS();
    
    // Set presentation properties
    pres.author = 'AI Presentation Generator';
    pres.company = 'AI Generated';
    pres.revision = '1';
    pres.subject = slidesStructure.presentationTitle || 'AI Generated Presentation';
    pres.title = slidesStructure.presentationTitle || 'AI Generated Presentation';

    // Define slide styles based on template theme
    const getSlideStyles = (theme) => {
      return {
        title: {
          x: 0.5, y: 0.7, w: 9, h: 1.2,
          fontSize: theme?.titleFontSize || 32,
          fontFace: theme?.titleFont || 'Arial',
          color: theme?.titleColor || '1F4E79',
          bold: true,
          align: 'center'
        },
        content: {
          x: 0.8, y: 2.2, w: 8.5, h: 4.5,
          fontSize: theme?.contentFontSize || 18,
          fontFace: theme?.contentFont || 'Arial',
          color: theme?.contentColor || '404040',
          valign: 'top'
        },
        background: theme?.background || { color: 'FFFFFF' }
      };
    };

    const styles = getSlideStyles(templateTheme);
    
    // Generate NEW slides using template styling
    slidesStructure.slides.forEach((slideData, index) => {
      const slide = pres.addSlide();
      
      console.log(`Creating new slide ${index + 1}: ${slideData.title}`);
      
      // Apply template background if available
      if (styles.background) {
        slide.background = styles.background;
      }
      
      // Add title with template styling
      slide.addText(slideData.title || `Slide ${index + 1}`, styles.title);
      
      // Add content with template styling
      if (slideData.content) {
        if (Array.isArray(slideData.content)) {
          // Handle bullet points
          const bulletPoints = slideData.content.map(point => ({
            text: point,
            options: { 
              bullet: true,
              fontSize: styles.content.fontSize,
              fontFace: styles.content.fontFace,
              color: styles.content.color
            }
          }));
          slide.addText(bulletPoints, {
            x: styles.content.x,
            y: styles.content.y,
            w: styles.content.w,
            h: styles.content.h,
            valign: styles.content.valign
          });
        } else {
          slide.addText(slideData.content, styles.content);
        }
      }
      
      // Add slide number
      slide.addText(`${index + 1}`, {
        x: 9.2, y: 6.8, w: 0.6, h: 0.3,
        fontSize: 12, fontFace: 'Arial', color: '666666',
        align: 'center'
      });
    });
    
    console.log(`Created ${slidesStructure.slides.length} new slides using template styling`);
    
    // Generate and return buffer
    return await pres.write({ outputType: 'nodebuffer' });
    
  } catch (error) {
    console.error('Error creating slides with template style:', error);
    return await generateNewPresentation(slidesStructure);
  }
}

// Extract theme information from template
async function extractTemplateTheme(templateBuffer) {
  try {
    const zip = new PizZip(templateBuffer);
    const theme = {
      titleFont: 'Arial',
      titleFontSize: 32,
      titleColor: '1F4E79',
      contentFont: 'Arial', 
      contentFontSize: 18,
      contentColor: '404040',
      background: { color: 'FFFFFF' }
    };
    
    // Try to extract theme colors from theme1.xml
    try {
      const themeFile = zip.files['ppt/theme/theme1.xml'];
      if (themeFile) {
        const themeXml = themeFile.asText();
        
        // Extract colors using simple regex
        const colorPatterns = {
          accent1: /<a:accent1[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:accent1>/i,
          accent2: /<a:accent2[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:accent2>/i,
          dk1: /<a:dk1[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"[\s\S]*?<\/a:dk1>/i
        };
        
        const accent1Match = themeXml.match(colorPatterns.accent1);
        const dk1Match = themeXml.match(colorPatterns.dk1);
        
        if (accent1Match) theme.titleColor = accent1Match[1];
        if (dk1Match) theme.contentColor = dk1Match[1];
        
        console.log(`Extracted colors - Title: ${theme.titleColor}, Content: ${theme.contentColor}`);
      }
    } catch (error) {
      console.log('Could not extract theme colors, using defaults');
    }
    
    // Try to extract fonts from first slide
    try {
      const firstSlide = zip.files['ppt/slides/slide1.xml'];
      if (firstSlide) {
        const slideXml = firstSlide.asText();
        
        // Look for font information
        const fontMatch = slideXml.match(/<a:latin typeface="([^"]*)"/);
        if (fontMatch) {
          theme.titleFont = fontMatch[1];
          theme.contentFont = fontMatch[1];
          console.log(`Extracted font: ${fontMatch[1]}`);
        }
      }
    } catch (error) {
      console.log('Could not extract fonts, using defaults');
    }
    
    return theme;
    
  } catch (error) {
    console.error('Error extracting template theme:', error);
    return null;
  }
}

// Helper function to escape XML characters
function escapeXml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Generate new presentation without template using pptxgenjs
async function generateNewPresentation(slidesStructure) {
  const pres = new PptxGenJS();
  
  // Set presentation properties
  pres.author = 'AI Presentation Generator';
  pres.company = 'AI Generated';
  pres.revision = '1';
  pres.subject = slidesStructure.presentationTitle || 'AI Generated Presentation';
  pres.title = slidesStructure.presentationTitle || 'AI Generated Presentation';

  // Generate slides with default styling
  slidesStructure.slides.forEach((slideData, index) => {
    const slide = pres.addSlide();
    
    // Add title
    slide.addText(slideData.title || `Slide ${index + 1}`, {
      x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 28, fontFace: 'Arial',
      color: '1F497D', bold: true, align: 'center'
    });
    
    // Add content
    if (slideData.content) {
      const contentStyle = {
        x: 0.5, y: 2, w: 9, h: 4, fontSize: 18, fontFace: 'Arial',
        color: '444444', valign: 'top'
      };
      
      if (Array.isArray(slideData.content)) {
        const bulletPoints = slideData.content.map(point => ({
          text: point,
          options: { bullet: true }
        }));
        slide.addText(bulletPoints, contentStyle);
      } else {
        slide.addText(slideData.content, contentStyle);
      }
    }
  });
  
  // Generate and return buffer
  return await pres.write({ outputType: 'nodebuffer' });
}

// Create the analysis prompt
function createAnalysisPrompt(text, guidance) {
  return `Analyze the following text and break it into a PowerPoint presentation structure. ${guidance ? `Context: ${guidance}` : ''}

Create 5-12 slides based on the content. For each slide, provide:
1. A clear, concise title (max 10 words)
2. Key content points (2-5 bullet points or 1-2 short paragraphs)
3. Slide type (title, content, summary, etc.)

Text to analyze:
${text}

Please respond in the following JSON format:
{
  "slides": [
    {
      "title": "Slide Title",
      "content": ["Bullet point 1", "Bullet point 2", "Bullet point 3"],
      "type": "content",
      "notes": "Optional speaker notes"
    }
  ],
  "totalSlides": 0,
  "presentationTitle": "Overall Presentation Title"
}

IMPORTANT: Return ONLY the JSON object, no additional text or explanation.`;
}

// Parse AI response to extract JSON
function parseAIResponse(aiResponse) {
  try {
    // Try to find JSON in the response
    const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.error('No JSON found in AI response');
      return null;
    }

    const parsed = JSON.parse(jsonMatch[0]);
    
    // Validate structure
    if (!parsed.slides || !Array.isArray(parsed.slides)) {
      console.error('Invalid slide structure in response');
      return null;
    }

    // Set totalSlides if not provided
    if (!parsed.totalSlides) {
      parsed.totalSlides = parsed.slides.length;
    }

    // Set default presentation title if not provided
    if (!parsed.presentationTitle) {
      parsed.presentationTitle = 'Generated Presentation';
    }

    return parsed;

  } catch (error) {
    console.error('Error parsing AI response:', error);
    return null;
  }
}

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({ 
    error: 'Internal server error',
    details: error.message 
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Endpoint not found' });
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 AI Presentation Generator Backend running on port ${PORT}`);
  console.log(`📱 Frontend: http://localhost:${PORT}`);
  console.log(`🔗 API Health: http://localhost:${PORT}/api/health`);
});

module.exports = app;
