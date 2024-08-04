/* global Word console */

// Function to search for and highlight words
export async function searchForWord(word: string): Promise<string[]> {
  if (!word || word.length > 100) {
    console.error("Search string is invalid or too long");
    return [];
  }

  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const searchResults = body.search(word, { matchCase: false, matchWholeWord: true });
      context.load(searchResults, "text");

      await context.sync();

      searchResults.items.forEach((item) => {
        item.font.highlightColor = "yellow";
      });

      await context.sync();

      return searchResults.items.map((item) => item.text);
    });
  } catch (error) {
    console.error("Error searching for word: " + error);
    return [];
  }
}

// Function to clear all highlights
export async function clearHighlights(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      context.load(paragraphs);
      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        paragraph.font.highlightColor = null;
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error clearing highlights: " + error);
  }
}

// Function to extract text with positions
export async function extractTextWithPositions(): Promise<{ text: string; startPosition: number }[]> {
  try {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, "text");
      await context.sync();

      let cumulativeLength = 0;
      const results = paragraphs.items.map((paragraph) => {
        const result = {
          text: paragraph.text,
          startPosition: cumulativeLength,
        };
        cumulativeLength += paragraph.text.length + 1; // Assuming a newline character
        return result;
      });

      return results;
    });
  } catch (error) {
    console.error("Error extracting text with positions: " + error);
    return [];
  }
}

// Function to get positions of each word
export async function getWordPositions(): Promise<{ word: string; from: number; to: number; isRed: boolean }[]> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      context.load(paragraphs, "text"); // Load text property of paragraphs
      await context.sync();

      let cumulativeLength = 0;
      const results = [];

      for (const paragraph of paragraphs.items) {
        const words = paragraph.text.split(/\s+/);
        for (const word of words) {
          const startPosition = cumulativeLength;
          const endPosition = startPosition + word.length - 1;

          results.push({
            word,
            from: startPosition,
            to: endPosition,
            isRed: false, // Example property
          });

          // Condition to check and replace the word at specific positions
          if (startPosition === 13 && endPosition === 16) {
            const wordRange = paragraph.getRange("Start").expandTo(paragraph.getRange("End"));
            context.load(wordRange); // Prepare to load the range to be replaced, operation is queued but not executed yet
            // Replace the text in the specified range
            wordRange.insertText("replacementWord", "Replace");
            console.log("Scheduled replacement for word at positions from 13 to 16");
          }

          cumulativeLength += word.length + 1; // Adjusting for spaces
        }
        cumulativeLength++; // Adjusting for newline characters
      }

      // Perform a single sync after all operations are queued
      await context.sync();
      console.log("Word Positions:", results);
      return results;
    });
  } catch (error) {
    console.error("Error getting and replacing word positions: " + error);
    return [];
  }
}

// Function to add a dropdown to words in positions from 10 to 15
export async function addDropdownToWordsInRange(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      context.load(paragraphs, "text");
      await context.sync();

      let wordCounter = 0;

      for (const paragraph of paragraphs.items) {
        const words = paragraph.text.split(/\s+/);
        for (const word of words) {
          if (word && wordCounter >= 10 && wordCounter <= 15) {
            const range = paragraph
              .getRange(Word.RangeLocation.start)
              .expandTo(paragraph.getRange(Word.RangeLocation.start).split([word])[1]);

            const dropdownControl = range.insertContentControl();
            dropdownControl.tag = `dropdown_${wordCounter}`;
            dropdownControl.title = `Dropdown for ${word}`;
            dropdownControl.insertText(word, Word.InsertLocation.replace);

            // Add dropdown items
            dropdownControl.insertParagraph("Option 1", Word.InsertLocation.end);
            dropdownControl.insertParagraph("Option 2", Word.InsertLocation.end);
            dropdownControl.insertParagraph("Option 3", Word.InsertLocation.end);

            wordCounter++;
          } else {
            wordCounter++;
          }
        }
      }

      await context.sync();
      console.log("Dropdown added to words in positions from 10 to 15 in the document.");
    });
  } catch (error) {
    console.error("Error adding dropdown: " + error);
  }
}

export async function replaceTextInRange(startPos: number, endPos: number, replacementText: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection(); // Get the current selection
      range.load("text");
      await context.sync();

      // Suppose startPos and endPos refer to the offsets within the currently selected range.
      // Adjust this approach if you need to select a specific range programmatically.
      const textToReplace = range.text.substring(startPos, endPos + 1);
      const replacedText = range.text.replace(textToReplace, replacementText);

      // Clear the current selection and replace it with the modified text
      range.clear();
      range.insertText(replacedText, "Replace");

      await context.sync();
    });
  } catch (error) {
    console.error("Error replacing text: " + error);
  }
}

export async function getWordPositionsAndReplace(): Promise<
  { word: string; from: number; to: number; isRed: boolean }[]
> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      context.load(paragraphs, "text"); // Load the text property of paragraphs
      await context.sync();

      let cumulativeLength = 0;
      const results = [];
      const changes = [];

      for (const paragraph of paragraphs.items) {
        const text = paragraph.text;

        if (text.length + cumulativeLength >= 13 && cumulativeLength <= 16) {
          const relativeStart = 13 - cumulativeLength;
          const relativeEnd = 16 - cumulativeLength;

          const prefix = text.slice(0, relativeStart);
          const suffix = text.slice(relativeEnd + 1);
          const newText = prefix + "replacementWord" + suffix;

          changes.push({ paragraph, newText });
        }

        for (const word of text.split(/\s+/)) {
          const startPosition = cumulativeLength;
          const endPosition = startPosition + word.length - 1;

          results.push({
            word,
            from: startPosition,
            to: endPosition,
            isRed: false,
          });

          cumulativeLength += word.length + 1; // Adjust for spaces
        }
        cumulativeLength++; // Adjust for new lines
      }

      // Perform the text replacements
      for (const change of changes) {
        change.paragraph.insertText(change.newText, Word.InsertLocation.replace);
      }

      await context.sync(); // Final synchronization after all operations
      console.log("Word Positions:", results);
      return results;
    });
  } catch (error) {
    console.error("Error getting and replacing word positions: " + error);
    return [];
  }
}
