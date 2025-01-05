import * as React from "react";
import { useState } from "react";
import { Button, Checkbox, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import PropTypes from "prop-types";
import images from "./images.json";
import targetJson from "./target_conditions.json";
import bulletJson from "./bullets.json";

const useStyles = makeStyles({
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "starts",
  },
  Title: {
    alignSelf: "center",
  },
  HR: {
    width: "100%"
  },
  P: {
    margin: "5px 0px 0px 0px",
    fontSize: "20px"
  },
  List: {
    display: "flex",
    flexDirection: "column",
    alignItems: "starts",
  }
});

const BulletInsertion = (props) => {
  const bullets = bulletJson.bullets;
  const [selectedBullets, setSelectedBullets] = useState(bullets.map(bullet => bullet.items.map(item => false)));

  const handleBulletInsertion = async () => {
    let bulletsToInsert = bullets.map((bullet, i) => ({...bullet, items: bullet.items.filter((item, j) => selectedBullets[i][j])}));
    bulletsToInsert = bulletsToInsert.filter(bullet => bullet.items.length > 0);
    await props.insertBullet(bulletsToInsert);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <p className={styles.Title}>Add Bulletpoints</p>
      {bullets.map((bullet, i) => (
        <div key={"bullet-" + i}>
          <p>{bullet.title}</p>
          <div className={styles.List}>
          {bullet.items.map((item, j) => (
            <Checkbox label={item} key={"bullet-" + i + "-key-" + j} checked={selectedBullets[i][j]}
                onChange={() => {
                  const newSelectedBullets = [...selectedBullets];
                  newSelectedBullets[i][j] = !selectedBullets[i][j];
                  setSelectedBullets(newSelectedBullets);
                }}
                />
          ))}
          </div>
        </div>
      ))}

      <Button className={styles.Title} appearance="primary" disabled={false} size="normal" onClick={handleBulletInsertion}>
        Insert Bullet List
      </Button>
      <hr className={styles.HR}></hr>
    </div>
  );
};

BulletInsertion.propTypes = {
  insertBullet: PropTypes.func.isRequired,
};

export default BulletInsertion;
