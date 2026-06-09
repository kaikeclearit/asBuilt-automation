import os
import re
import time
import yaml
import logging
import winpty
from datetime import datetime

# ---------------------------------------------------------------------------
# DEBUG LOGGER SETUP
# Writes to:  <project_root>/logs/vendor_generator_YYYYMMDD_HHMMSS.log
# AND streams to console/Streamlit stdout simultaneously
# ---------------------------------------------------------------------------

def _setup_logger() -> logging.Logger:
    log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"vendor_generator_{timestamp}.log")

    logger = logging.getLogger("vendor_generator")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()  # prevent duplicate handlers on Streamlit reruns

    # File handler — full detail
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    # Console/stdout handler — same detail
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)

    logger.info(f"=== Session started — log file: {log_path} ===")
    return logger, log_path


# ---------------------------------------------------------------------------
# MAIN FUNCTION
# ---------------------------------------------------------------------------

def run_silent_nutanix_generation(tool_dir, customer_name, target_ip, username, password):
    log, log_path = _setup_logger()

    log.info(f"tool_dir     : {tool_dir}")
    log.info(f"customer_name: {customer_name}")
    log.info(f"target_ip    : {target_ip}")
    log.info(f"username     : {username}")
    log.info(f"password     : {'*' * len(password)}")

    # ------------------------------------------------------------------
    # 1. Validate exe exists before doing anything
    # ------------------------------------------------------------------
    exe_path = os.path.normpath(os.path.join(tool_dir, "run_as_built.exe"))
    tool_dir  = os.path.normpath(tool_dir)

    log.info(f"exe_path     : {exe_path}")
    log.info(f"exe exists   : {os.path.exists(exe_path)}")
    log.info(f"exe executable: {os.access(exe_path, os.X_OK)}")

    if not os.path.exists(exe_path):
        raise FileNotFoundError(f"run_as_built.exe not found at: {exe_path}")

    # ------------------------------------------------------------------
    # 2. Write YAML config
    # ------------------------------------------------------------------
    yaml_config_path = os.path.join(tool_dir, "user_file_input.yml")
    config_payload = {"pc": [target_ip]}
    with open(yaml_config_path, "w") as f:
        yaml.dump(config_payload, f, default_flow_style=False)
    log.info(f"YAML config written to: {yaml_config_path}")
    log.debug(f"YAML content: {config_payload}")

    # ------------------------------------------------------------------
    # 3. Change to tool directory (binary resolves paths relative to cwd)
    # ------------------------------------------------------------------
    original_cwd = os.getcwd()
    os.chdir(tool_dir)
    log.info(f"cwd changed to: {os.getcwd()}")

    pty = None
    try:
        # --------------------------------------------------------------
        # 4. Spawn PTY
        # --------------------------------------------------------------
        command = [exe_path, "-c", customer_name, "-s", "cluster", "-i", "f"]
        log.info(f"Spawning PTY command: {command}")
        pty = winpty.PtyProcess.spawn(command)
        log.info("PTY spawned successfully")

        steps_sent = {
            "volumes"  : False,
            "cloud"    : False,
            "file_path": False,
            "user"     : False,
            "pass"     : False,
        }

        buffer  = ""
        timeout = 120
        start   = time.time()
        chunk_count = 0

        # --------------------------------------------------------------
        # 5. PTY read loop
        # --------------------------------------------------------------
        while pty.isalive():
            elapsed = time.time() - start
            if elapsed > timeout:
                log.error(f"Timeout after {timeout}s — killing PTY")
                pty.terminate()
                raise TimeoutError(f"Nutanix binary timed out after {timeout}s")

            try:
                chunk = pty.read(1024)
            except EOFError:
                log.info("PTY EOF — process finished")
                break

            if not chunk:
                time.sleep(0.05)
                continue

            chunk_count += 1
            buffer += chunk
            clean = _strip_ansi(buffer)

            # Log every raw chunk so you can see exactly what the binary sends
            log.debug(f"[CHUNK #{chunk_count} | +{elapsed:.1f}s] RAW: {repr(chunk)}")
            log.debug(f"[CHUNK #{chunk_count}] CLEAN: {repr(clean[-300:])}")  # last 300 chars

            # ----------------------------------------------------------
            # 6. Step matching + responses
            # ----------------------------------------------------------
            if not steps_sent["volumes"] and "volumes in NC2" in clean and "[y/n]:" in clean:
                log.info(">>> MATCH: volumes question — sending 'n'")
                pty.write("n\r")
                steps_sent["volumes"] = True
                buffer = ""

            elif not steps_sent["cloud"] and "AWS/Azure in NC2" in clean and "[y/n]:" in clean:
                log.info(">>> MATCH: cloud question — sending 'n'")
                pty.write("n\r")
                steps_sent["cloud"] = True
                buffer = ""

            elif not steps_sent["file_path"] and "[user_file_input.yml]:" in clean:
                log.info(">>> MATCH: file path prompt — sending ENTER")
                pty.write("\r")
                steps_sent["file_path"] = True
                buffer = ""

            elif not steps_sent["user"] and "username for the PC" in clean and ":" in clean:
                log.info(f">>> MATCH: username prompt — sending '{username}'")
                pty.write(f"{username}\r")
                steps_sent["user"] = True
                buffer = ""

            elif not steps_sent["pass"] and "Password for" in clean:
                log.info(">>> MATCH: password prompt — sending password")
                pty.write(f"{password}\r")
                steps_sent["pass"] = True
                buffer = ""

            elif steps_sent["pass"] and _generation_complete(clean):
                log.info(">>> MATCH: generation complete signal detected — exiting loop")
                break

        log.info(f"PTY loop ended — steps completed: {steps_sent}")
        log.info(f"Total chunks read: {chunk_count}")
        log.info(f"Total elapsed: {time.time() - start:.1f}s")

        # --------------------------------------------------------------
        # 7. Find the generated output file
        # --------------------------------------------------------------
        output_dir = os.path.join(tool_dir, "as_built_docs", "cluster")
        log.info(f"Checking output dir: {output_dir}")

        if os.path.exists(output_dir):
            files = os.listdir(output_dir)
            log.info(f"Files in output dir: {files}")
        else:
            log.warning(f"Output dir does not exist: {output_dir}")

        return output_dir, log_path

    except Exception as e:
        log.exception(f"FATAL ERROR in PTY pipeline: {e}")
        raise

    finally:
        os.chdir(original_cwd)
        log.info(f"cwd restored to: {original_cwd}")
        if pty is not None:
            try:
                pty.terminate(force=True)
                log.info("PTY terminated cleanly")
            except Exception as ex:
                log.warning(f"PTY terminate warning: {ex}")
        log.info("=== Session ended ===")


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def _generation_complete(text: str) -> bool:
    signals = ["successfully", "as_built_docs", "complete", "generated", "Finished"]
    return any(s.lower() in text.lower() for s in signals)


def _strip_ansi(text: str) -> str:
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
    return ansi_escape.sub("", text)