import * as React from "react";
import PropTypes from "prop-types";
import { tokens, makeStyles } from "@fluentui/react-components";
import { useState, useEffect } from "react";

const useStyles = makeStyles({
  list: {
    marginTop: "0px",
    marginLeft: "0px"
  },
  listItem: {
    padding: '0px 0px 10px 10px',
    display: "flex",
    cursor: "pointer",
    textAlign: "left",
    alignItems: "center"
  },
  icon: {
    marginRight: "10px",
  },
  itemText: {
    fontSize: tokens.fontSizeBase300,
    fontColor: tokens.colorNeutralBackgroundStatic,
    position: "relative",
    left: "-30px",
    fontSize: "15px"
  },
  welcome__main: {
    width: "50%",
    display: "flex",
    flexDirection: "column",
    alignItems: "left",
  },
  message: {
    fontSize: tokens.fontSizeBase500,
    fontColor: tokens.colorNeutralBackgroundStatic,
    fontWeight: tokens.fontWeightRegular,
    paddingLeft: "0px",
    paddingRight: "0px",
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const HeroList = (props) => {
  const [text, setText] = useState("Some text.");
  const { items, message } = props;
  const styles = useStyles();
  const [template, setTemplate] = useState('');

  const handleTextInsertion = async (url) => {
    const api = "https://m3windsignature-bucabmeuhxaafda3.uksouth-01.azurewebsites.net/api/Signatures/signatures?signatureURL=" + url;
    fetch(api)
      .then((response) => response.json())
      .then((data) => {
        console.log(data);
        setTemplate(data.result);
        props.insertText(template);
      })
      .catch((err) => {
        console.log(err.message);
      });
  };

  const listItems = items.map((item, index) => (
    <li onClick={() => handleTextInsertion(item.url)} className={styles.listItem} key={index}>
      <span className={styles.icon}>{item.icon}</span>
      <span className={styles.itemText}>{item.primaryText}</span>
    </li>
  ));
  return (
    <ul className={styles.list}>{listItems}</ul>
  );
};

HeroList.propTypes = {
  items: PropTypes.arrayOf(
    PropTypes.shape({
      icon: PropTypes.element, // Optional
      primaryText: PropTypes.string.isRequired,
      url: PropTypes.string
    })
  ).isRequired,
  message: PropTypes.string.isRequired,
  insertText: PropTypes.func.isRequired,
};

export default HeroList;