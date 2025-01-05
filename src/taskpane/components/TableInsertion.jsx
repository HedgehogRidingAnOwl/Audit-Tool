import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import PropTypes from "prop-types";
import images from "./images.json";
import targetJson from "./target_conditions.json";

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
  const targetconditions = targetJson.targetconditions;
  const [target, setTarget] = useState(targetconditions[0]);

  const handleTableInsertion = async () => {
    await props.insertTable(target.content,"","","","", images.red);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <p className={styles.P}>Add Target Table</p>
      <select value={target.title} className={styles.dropDown} onChange={(e) => setTarget(targetconditions.find((condition) => condition.title === e.target.value))}>
        {targetconditions.map((condition) => (
          <option key={condition.title} value={condition.title}>{condition.title}</option>
        ))}
      </select>
      <Button appearance="primary" disabled={false} size="normal" onClick={handleTableInsertion}>
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
