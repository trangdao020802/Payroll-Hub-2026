import React from "react";
import { motion, AnimatePresence } from "motion/react";
import { HoldAddDashboard } from "./components/HoldAddDashboard";

export function HoldDashboardPage(): React.ReactElement {
  return (
    <div className="trial-balance-page-content flex-1 flex flex-col min-h-0 relative overflow-hidden" style={{ padding: "24px 24px" }}>
      <AnimatePresence initial={false}>
        <motion.div
          key="hold-dashboard-main"
          initial={{ y: 6, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.14, ease: "easeOut" }}
          style={{ padding: "12px" }}
          className="w-full h-full flex flex-col min-h-0 bg-transparent items-center overflow-hidden"
        >
          <HoldAddDashboard />
        </motion.div>
      </AnimatePresence>
    </div>
  );
}
