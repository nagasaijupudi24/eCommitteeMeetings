import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";

const DateTime: React.FC = () => {
  const [currentDate, setCurrentDate] = useState(new Date());

  useEffect(() => {
    const timerID = setInterval(() => setCurrentDate(new Date()), 1000);
    return () => clearInterval(timerID);
  }, []);

  const formattedDate: string = `${currentDate.getDate()}-${
    currentDate.getMonth() + 1
  }-${currentDate.getFullYear()} ${currentDate.getHours()}:${currentDate.getMinutes()}:${currentDate.getSeconds()}`;

  return (
    <p style={{ fontSize: "1rem", margin: 0 }} className={styles.titleDate}>
      Date: {formattedDate}
    </p>
  );
};

export default DateTime;
