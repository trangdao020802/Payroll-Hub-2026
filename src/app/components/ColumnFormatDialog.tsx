import React, { useState } from "react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "./ui/dialog";
import { Button } from "./ui/button";
import { Label } from "./ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./ui/select";

interface ColumnFormatDialogProps {
  isOpen: boolean;
  onClose: () => void;
  colKey: string;
  initialFormat: { alignment?: "left" | "center" | "right" };
  onSave: (format: { alignment?: "left" | "center" | "right" }) => void;
}

export const ColumnFormatDialog: React.FC<ColumnFormatDialogProps> = ({
  isOpen,
  onClose,
  colKey,
  initialFormat,
  onSave,
}) => {
  const [alignment, setAlignment] = useState<"left" | "center" | "right">(
    initialFormat.alignment || "left",
  );

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Định dạng cột: {colKey}</DialogTitle>
        </DialogHeader>
        <div className="grid gap-4 py-4">
          <div className="grid grid-cols-4 items-center gap-4">
            <Label htmlFor="alignment" className="text-right">
              Căn chỉnh
            </Label>
            <Select
              value={alignment}
              onValueChange={(val) => {
                if (val === "left" || val === "center" || val === "right") setAlignment(val);
              }}
            >
              <SelectTrigger className="col-span-3">
                <SelectValue placeholder="Chọn căn chỉnh" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="left">Trái</SelectItem>
                <SelectItem value="center">Giữa</SelectItem>
                <SelectItem value="right">Phải</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={onClose}>
            Hủy
          </Button>
          <Button
            onClick={() => {
              onSave({ alignment });
              onClose();
            }}
          >
            Lưu
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
};
