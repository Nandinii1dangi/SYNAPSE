import cv2
import os
import time
from attendance_handler import save_student_to_registry


def enroll_new_student():
    print("\n=========================================")
    print("      SYNAPSE BIOMETRIC ENROLLMENT       ")
    print("=========================================\n")

    student_name = input("Enter the name of student: ").strip().upper()
    student_id = input("Enter the student id: ").strip().upper()

    if not student_name or not student_id:
        print("[CRITICAL ERROR] Inputs cannot be blank. Enrollment aborted.")
        return

    # Save to JSON registry file
    save_student_to_registry(student_name, student_id)

    base_folder = "known_faces"
    student_folder = os.path.join(base_folder, student_name)
    os.makedirs(student_folder, exist_ok=True)

    print(f"\n[SYSTEM] Ready to capture face frames for: {student_name}")
    print("[SYSTEM] Look directly at the webcam. Starting capture loop in 3 seconds...")
    time.sleep(3)

    camera = cv2.VideoCapture(0)
    if not camera.isOpened():
        print("[CRITICAL ERROR] Could not open webcam.")
        return

    sample_count = 0
    max_samples = 20

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

    print("[CAPTURE STARTING] Taking snapshot matrices live...")

    # Track time between captures without freezing cv2.imshow
    last_capture_time = 0
    capture_delay = 0.2  # Seconds between image saves

    while sample_count < max_samples:
        success, frame = camera.read()
        if not success:
            print("[ERROR] Failed to grab frame from webcam.")
            break

        # Flip frame immediately so processing and drawing match perfectly
        frame = cv2.flip(frame, 1)
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(100, 100))

        current_time = time.time()

        for (x, y, w, h) in faces:
            # Only save a sample if the delay window has passed
            if current_time - last_capture_time >= capture_delay:
                sample_count += 1
                img_name = f"{student_name}_sample_{sample_count}.jpg"
                img_path = os.path.join(student_folder, img_name)

                # Crop correctly from the flipped frame matrix
                face_crop = frame[y:y + h, x:x + w]
                cv2.imwrite(img_path, face_crop)

                print(f" -> Captured Sample {sample_count}/{max_samples}")
                last_capture_time = current_time

            # Draw visual feedback box
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
            break  # Process only one face per frame

        cv2.imshow("SYNAPSE Face Sample Enrollment Pipeline", frame)

        # Break loop if 'q' key is pressed
        if cv2.waitKey(1) & 0xFF == ord('q'):
            print("[INFO] Enrollment interrupted by user.")
            break

    camera.release()
    cv2.destroyAllWindows()
    print(f"\n[SUCCESS] 🎓 Enrollment complete for {student_name}!")


if __name__ == "__main__":
    enroll_new_student()
