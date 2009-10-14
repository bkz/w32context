(defun w32-context-menu (filename)
  "Trigger shell context menu via external helper program."
  (start-process-shell-command "w32context" "*w32context*" "w32context" filename))

(defun w32-context-menu-dired-mouse-event (event)
  "Invoke the shell context menu in Do-What-I-Mean sense."
  (interactive "e")
  (mouse-set-point event)
  (w32-context-menu (dwim-get-current-full-path)))

(defun w32-context-menu-dwim ()
  "Invoke the shell context menu in Do-What-I-Mean sense."
  (interactive)
  (w32-context-menu (dwim-get-current-full-path)))

(defun dired-get-selected-item-full-path ()
  "Retrieve the full path for the currently selected file or directory."
  (let ((curr-item (dired-get-file-for-visit)))
    (progn
      (when (file-directory-p curr-item)
        (setq curr-item (concat curr-item "/")))
      (when(equal 'windows-nt system-type)
        (setq curr-item (dired-replace-in-string "/" "\\" curr-item)))
      curr-item)))

(defun dwim-get-current-full-path ()
  "If the current buffer is a dired buffer, retrieve the path of
  the selected item otherwise return the path of the buffer. If
  the current buffer does not have a path associated with return
  the default directory."
  (if (eq major-mode 'dired-mode)
      (dired-get-selected-item-full-path)
    (let ((path (buffer-file-name)))
      (when (null path)
        (setq path default-directory))
      (when(equal 'windows-nt system-type)
        (setq path (dired-replace-in-string "/" "\\" path)))
      path)))


;; Bind right-click mouse button and a keyboard hotkey to trigger the Explorer
;; shell context menu from within Emacs dired and normal file buffers.
(define-key dired-mode-map [e)]      'w32-context-menu-dwim)
(define-key dired-mode-map [mouse-3] 'w32-context-menu-dired-mouse-event))
