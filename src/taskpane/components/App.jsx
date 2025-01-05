import * as React from "react";
import PropTypes from "prop-types";
import TableInsertion from "./TableInsertion";
import { makeStyles } from "@fluentui/react-components";
import { insertTable, insertBullet } from "../taskpane";
import BulletInsertion from "./BulletInsertion";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "5px",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.

  return (
    <div className={styles.root}>
      <TableInsertion insertTable={insertTable} />
      <BulletInsertion insertBullet={insertBullet} />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
