import os
import random
import asyncio
import aiohttp
import time
import platform
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from pydub import AudioSegment
import speech_recognition as sr


class RecaptchaSolver:
    def __init__(self, driver, debug_mode=False):
        self.driver = driver
        self.debug_mode = debug_mode
        self.is_windows_10 = (
            platform.system() == "Windows" and platform.release().startswith("10")
        )

        if self.debug_mode:
            print(
                f"RecaptchaSolver inicializado no {'Windows 10' if self.is_windows_10 else platform.system() + ' ' + platform.release()}"
            )

    async def download_audio(self, url, path):
        async with aiohttp.ClientSession() as session:
            async with session.get(url) as response:
                with open(path, "wb") as f:
                    f.write(await response.read())
        print("Downloaded audio asynchronously.")

    def solveCaptcha(self):
        if self.debug_mode:
            print("Iniciando solução de CAPTCHA...")

        # Detectar o tipo de CAPTCHA presente na página
        captcha_type = self._detect_captcha_type()

        if self.debug_mode:
            print(f"Tipo de CAPTCHA detectado: {captcha_type}")

        # Ajustar timeouts para o Windows 10
        wait_time = 5
        if self.is_windows_10:
            wait_time = 10

        try:
            # Switch to the CAPTCHA iframe
            iframe_inner = WebDriverWait(self.driver, wait_time).until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.XPATH, "//iframe[contains(@title, 'reCAPTCHA')]")
                )
            )

            # Click on the CAPTCHA box
            WebDriverWait(self.driver, wait_time).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-anchor"))
            ).click()

            # Check if the CAPTCHA is solved
            time.sleep(1)  # Allow some time for the state to update
            if self.isSolved():
                print("CAPTCHA solved by clicking.")
                self.driver.switch_to.default_content()  # Switch back to main content
                return

            # If not solved, attempt audio CAPTCHA solving
            self.solveAudioCaptcha()

        except Exception as e:
            print(f"An error occurred while solving CAPTCHA: {e}")
            self.driver.switch_to.default_content()  # Ensure we switch back in case of error
            raise

    def _detect_captcha_type(self):
        """Detectar o tipo de CAPTCHA presente na página"""
        try:
            # Verificar reCAPTCHA v2
            recaptcha_frames = self.driver.find_elements(
                By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"
            )
            if recaptcha_frames:
                return "reCAPTCHA v2"

            # Verificar hCAPTCHA
            hcaptcha_frames = self.driver.find_elements(
                By.CSS_SELECTOR, "iframe[src*='hcaptcha.com']"
            )
            if hcaptcha_frames:
                return "hCAPTCHA"

            # Verificar CAPTCHA simples (apenas imagem e campo de texto)
            captcha_images = self.driver.find_elements(
                By.CSS_SELECTOR, "img[src*='captcha']"
            )
            if captcha_images:
                return "Simple Image CAPTCHA"

            return "Unknown"
        except Exception as e:
            if self.debug_mode:
                print(f"Erro ao detectar tipo de CAPTCHA: {str(e)}")
            return "Detection Error"

    def clickRefreshButton(self):
        """Tenta clicar no botão de renovar CAPTCHA"""
        try:
            # Procurar pelo botão de atualização que geralmente tem ID recaptcha-reload-button
            refresh_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-reload-button"))
            )
            print(
                "Botão de atualização de CAPTCHA encontrado. Clicando para obter um novo CAPTCHA..."
            )
            refresh_button.click()
            time.sleep(1.5)  # Aguardar o carregamento do novo CAPTCHA
            return True
        except Exception as e:
            print(f"Não foi possível encontrar ou clicar no botão de atualização: {e}")

            # Tentar localizar por XPath alternativo
            try:
                refresh_xpath = "//button[contains(@title, 'Get a new challenge') or contains(@class, 'reload')]"
                refresh_button = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, refresh_xpath))
                )
                print(
                    "Botão de atualização encontrado por XPath alternativo. Clicando..."
                )
                refresh_button.click()
                time.sleep(1.5)
                return True
            except:
                print(
                    "Nenhum botão de atualização encontrado por métodos alternativos."
                )
                return False

    def solveAudioCaptcha(self):
        try:
            self.driver.switch_to.default_content()

            # Switch to the audio CAPTCHA iframe
            iframe_audio = WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it(
                    (
                        By.XPATH,
                        '//iframe[@title="recaptcha challenge expires in two minutes"]',
                    )
                )
            )

            # Click on the audio button - support both English and Portuguese selectors
            try:
                audio_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "recaptcha-audio-button"))
                )
            except:
                # Tenta localizar botão de áudio em português ou por outras características
                audio_xpath = "//button[contains(@title, 'áudio') or contains(@title, 'audio') or contains(@class, 'audio-button')]"
                audio_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, audio_xpath))
                )
                print("Botão de áudio localizado via XPath alternativo")

            audio_button.click()

            # Tentar resolver até 3 CAPTCHAs diferentes
            for attempt in range(1, 4):
                try:
                    print(f"\nTentativa {attempt} com CAPTCHA de áudio...")

                    # Get the audio source URL
                    audio_source = (
                        WebDriverWait(self.driver, 10)
                        .until(EC.presence_of_element_located((By.ID, "audio-source")))
                        .get_attribute("src")
                    )
                    print(f"Audio source URL: {audio_source}")

                    # Download the audio to the temp folder asynchronously
                    temp_dir = os.getenv("TEMP") if os.name == "nt" else "/tmp/"
                    path_to_mp3 = os.path.normpath(
                        os.path.join(temp_dir, f"{random.randrange(1, 1000)}.mp3")
                    )
                    path_to_wav = os.path.normpath(
                        os.path.join(temp_dir, f"{random.randrange(1, 1000)}.wav")
                    )

                    asyncio.run(self.download_audio(audio_source, path_to_mp3))

                    # Convert mp3 to wav
                    sound = AudioSegment.from_mp3(path_to_mp3)
                    sound.export(path_to_wav, format="wav")
                    print("Converted MP3 to WAV.")

                    # Recognize the audio
                    recognizer = sr.Recognizer()
                    with sr.AudioFile(path_to_wav) as source:
                        audio = recognizer.record(source)
                    captcha_text = recognizer.recognize_google(audio).lower()
                    print(f"Recognized CAPTCHA text: {captcha_text}")

                    # Enter the CAPTCHA text
                    audio_response = WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.ID, "audio-response"))
                    )
                    audio_response.send_keys(captcha_text)
                    audio_response.send_keys(Keys.ENTER)
                    print("Entered and submitted CAPTCHA text.")

                    # Wait for CAPTCHA to be processed
                    time.sleep(1.2)  # Increase this if necessary

                    # Verify CAPTCHA is solved
                    if self.isSolved():
                        print("Audio CAPTCHA solved successfully.")
                        return True

                    # Se chegou aqui, o CAPTCHA não foi resolvido
                    print(f"Tentativa {attempt} falhou - resposta de áudio incorreta.")

                    # Se não for a última tentativa, clicar no botão de atualizar
                    if attempt < 3:
                        if not self.clickRefreshButton():
                            print("Não foi possível obter um novo CAPTCHA. Desistindo.")
                            break
                        print("Novo CAPTCHA carregado. Tentando novamente.")

                except Exception as e:
                    print(f"Erro durante a tentativa {attempt}: {e}")
                    # Se não for a última tentativa, tentar obter um novo CAPTCHA
                    if attempt < 3:
                        if not self.clickRefreshButton():
                            print("Não foi possível obter um novo CAPTCHA após erro.")
                            break
                        print("Novo CAPTCHA carregado após erro. Tentando novamente.")

            # Se chegou aqui, todas as tentativas falharam
            raise Exception(
                "Falha ao resolver CAPTCHAs de áudio após múltiplas tentativas"
            )

        except Exception as e:
            print(f"An error occurred while solving audio CAPTCHA: {e}")
            self.driver.switch_to.default_content()  # Ensure we switch back in case of error
            raise

        finally:
            # Always switch back to the main content
            self.driver.switch_to.default_content()

    def isSolved(self):
        try:
            # Switch back to the default content
            self.driver.switch_to.default_content()

            # Switch to the reCAPTCHA iframe
            iframe_check = self.driver.find_element(
                By.XPATH, "//iframe[contains(@title, 'reCAPTCHA')]"
            )
            self.driver.switch_to.frame(iframe_check)

            # Find the checkbox element and check its aria-checked attribute
            checkbox = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "recaptcha-anchor"))
            )
            aria_checked = checkbox.get_attribute("aria-checked")

            # Return True if the aria-checked attribute is "true" or the checkbox has the 'recaptcha-checkbox-checked' class
            return (
                aria_checked == "true"
                or "recaptcha-checkbox-checked" in checkbox.get_attribute("class")
            )

        except Exception as e:
            print(f"An error occurred while checking if CAPTCHA is solved: {e}")
            return False
