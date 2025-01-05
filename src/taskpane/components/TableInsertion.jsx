import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import PropTypes from "prop-types";
import images from "./images.json";

const useStyles = makeStyles({
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  dropDown: {
    margin: "15px 0px 20px 0px",
    fontSize: "20px"
  },
  HR: {
    width: "100%"
  },
  P: {
    margin: "5px 0px 0px 0px",
    fontSize: "20px"
  }
});

const TableInsertion = (props) => {
  const [text, setText] = useState("Some text.");

  const handleTableInsertion = async () => {
    await props.insertTable("","","","","", images.red);
  };

  const handleTextChange = async (event) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <p className={styles.P}>Add Target Table</p>
      <select value={1} className={styles.dropDown}>
        <option value={1}>First</option>
      </select>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTableInsertion}>
        Insert Table
      </Button>
      <hr className={styles.HR}></hr>
    </div>
  );
};

TableInsertion.propTypes = {
  insertTable: PropTypes.func.isRequired,
};

export default TableInsertion;
