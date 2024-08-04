/* global console */
import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import {
  searchForWord,
  clearHighlights,
  extractTextWithPositions,
  getWordPositions,
  addDropdownToWordsInRange,
  getWordPositionsAndReplace,
} from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "20px",
  },
  button: {
    marginTop: "20px",
    marginRight: "10px",
  },
  input: {
    marginRight: "10px",
  },
});

const App: React.FC<AppProps> = (props) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = React.useState("");
  const [searchResults, setSearchResults] = React.useState<string[]>([]);
  const [wordPositions, setWordPositions] = React.useState([]);
  const [textPositions, setTextPositions] = React.useState<{ text: string; startPosition: number }[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);

  const handleSearch = async () => {
    setIsLoading(true);
    const results = await searchForWord(searchTerm);
    setSearchResults(results);
    setIsLoading(false);
  };

  const handleClearHighlights = async () => {
    setIsLoading(true);
    await clearHighlights();
    setIsLoading(false);
  };

  const handleGetWordPositions = async () => {
    setIsLoading(true);
    const positions = await getWordPositions();
    setWordPositions(positions);
    setIsLoading(false);
    console.log("Word Positions:", positions); // Debugging log
  };

  const handleExtractTextWithPositions = async () => {
    setIsLoading(true);
    const positions = await extractTextWithPositions();
    setTextPositions(positions);
    setIsLoading(false);
  };

  const handleAddDropdownToWordsInRange = async () => {
    await addDropdownToWordsInRange();
  };

  const handleGetWordPositionsAndReplace = async () => {
    setIsLoading(true);
    const positions = await getWordPositionsAndReplace();
    setWordPositions(positions); // Assuming you want to update the same state
    setIsLoading(false);
    console.log("Updated Word Positions with Replacement:", positions);
  };

  return (
    <div className={styles.root}>
      <p>{props.title}</p>
      <input
        className={styles.input}
        type="text"
        value={searchTerm}
        onChange={(e) => setSearchTerm(e.target.value)}
        placeholder="Enter a word to search"
      />
      <button className={styles.button} onClick={handleSearch}>
        Search
      </button>
      <button className={styles.button} onClick={handleClearHighlights}>
        Clear Highlights
      </button>
      <button className={styles.button} onClick={handleGetWordPositions}>
        Get Word Positions
      </button>
      <button className={styles.button} onClick={handleExtractTextWithPositions}>
        Extract Text With Positions
      </button>
      <button className={styles.button} onClick={handleAddDropdownToWordsInRange}>
        Add Dropdown to Words in Range
      </button>
      <button className={styles.button} onClick={handleGetWordPositionsAndReplace}>
        Get and Replace Word Positions
      </button>
      {isLoading ? (
        <p>Loading...</p>
      ) : (
        <div>
          {searchResults.length > 0 && (
            <div>
              <h3>Search Results:</h3>
              {searchResults.map((result, index) => (
                <p key={index}>{result}</p>
              ))}
            </div>
          )}
          {wordPositions.length > 0 && (
            <div>
              <h3>Word Positions:</h3>
              {wordPositions.map((position, index) => (
                <div key={index}>
                  <p>
                    {position.word}: From {position.from} to {position.to}
                  </p>
                </div>
              ))}
            </div>
          )}
          {textPositions.length > 0 && (
            <div>
              <h3>Extracted Text and Positions:</h3>
              {textPositions.map((item, index) => (
                <div key={index}>
                  <p>Start Position: {item.startPosition}</p>
                  <p>{item.text}</p>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default App;
