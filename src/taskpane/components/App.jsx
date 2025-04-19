import * as React from "react";
import { useState, useEffect } from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Mail24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    maxHeight: "200vh",
  },
  innerDiv: {
    maxHeight: "100vh",
    maxWidth: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [signatures, setSignatures] = useState([]);
  const api = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Ribbons/ribbons";
  useEffect(() => {
    fetch(api)
      .then((response) => response.json())
      .then((data) => {
        console.log(data.result);
        setSignatures(data.result);
      })
      .catch((err) => {
        console.log(err.message);
      });
  }, []);

  const listItems = signatures.map((objSignature) => ({
    primaryText: objSignature.signature,
    url: objSignature.url,
    icon: <Mail24Regular />,
  }));

  return (
    <div className={styles.root}>
      <div className={styles.innerDiv}>
        <HeroList message="Select a signature" items={listItems} insertText={insertText} />
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
