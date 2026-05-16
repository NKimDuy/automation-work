"""
Pipeline Phân Loại Văn Bản Tiếng Việt
======================================
- Đầu vào : file CSV/Excel có 2 cột (text, label)
- Nhãn    : 1 đến 6
- Mô hình : TF-IDF + SVM (LinearSVC) — tối ưu cho tập nhỏ
- Output  : báo cáo accuracy, F1, confusion matrix + lưu model
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import joblib
import re
import os

from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    classification_report, confusion_matrix, accuracy_score
)
from sklearn.pipeline import Pipeline


# ─────────────────────────────────────────────
# 1. CẤU HÌNH — chỉnh tại đây
# ─────────────────────────────────────────────
FILE_PATH   = "./tieu_chi_lms_aug.xlsx"       # Đường dẫn file CSV hoặc Excel (.xlsx)
TEXT_COL    = "tieu_chi"           # Tên cột chứa cụm từ
LABEL_COL   = "nhan"          # Tên cột chứa nhãn (1–6)
TEST_SIZE   = 0.2              # 20% dùng để test
RANDOM_SEED = 42


# ─────────────────────────────────────────────
# 2. LOAD DỮ LIỆU
# ─────────────────────────────────────────────
def load_data(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[-1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding="utf-8-sig")

    print(f"✅ Đã load {len(df)} dòng từ '{path}'")
    print(f"   Phân phối nhãn:\n{df[LABEL_COL].value_counts().sort_index()}\n")
    return df


# ─────────────────────────────────────────────
# 3. TIỀN XỬ LÝ VĂN BẢN TIẾNG VIỆT
# ─────────────────────────────────────────────
def preprocess(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    # Xoá ký tự đặc biệt, giữ chữ cái và khoảng trắng
    text = re.sub(r"[^\w\sàáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ─────────────────────────────────────────────
# 4. XÂY DỰNG & HUẤN LUYỆN MÔ HÌNH
# ─────────────────────────────────────────────
def build_pipeline() -> Pipeline:
    return Pipeline([
        ("tfidf", TfidfVectorizer(
            analyzer="word",
            ngram_range=(1, 3),    # unigram + bigram + trigram
            max_features=10_000,
            sublinear_tf=True,     # giảm ảnh hưởng tần suất cao
            min_df=1,
        )),
        ("clf", LinearSVC(
            C=1.0,
            max_iter=2000,
            random_state=RANDOM_SEED
        ))
    ])


def train(df: pd.DataFrame):
    X = df[TEXT_COL].apply(preprocess)
    y = df[LABEL_COL].astype(int)

    # Cross-validation trên toàn bộ data
    model = build_pipeline()
    cv_scores = cross_val_score(model, X, y, cv=5, scoring="f1_macro")
    print(f"📊 Cross-validation F1 (5-fold): {cv_scores.mean():.3f} ± {cv_scores.std():.3f}\n")

    # Train/Test split
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=TEST_SIZE, random_state=RANDOM_SEED, stratify=y
    )

    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    # ── Báo cáo ──
    acc = accuracy_score(y_test, y_pred)
    print(f"🎯 Accuracy trên test set : {acc:.3f}")
    print("\n📋 Classification Report:")
    print(classification_report(y_test, y_pred, digits=3))

    # ── Confusion Matrix ──
    cm = confusion_matrix(y_test, y_pred)
    labels = sorted(y.unique())
    plt.figure(figsize=(7, 6))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Blues",
                xticklabels=labels, yticklabels=labels)
    plt.title("Confusion Matrix")
    plt.xlabel("Predicted Label")
    plt.ylabel("True Label")
    plt.tight_layout()
    # plt.savefig("confusion_matrix.png", dpi=150)
    plt.show()
    # print("💾 Đã lưu confusion matrix → confusion_matrix.png")

    # ── Lưu model ──
    joblib.dump(model, "model_vi_classification.pkl")
    print("💾 Đã lưu model → model_vi_classification.pkl")

    return model


# ─────────────────────────────────────────────
# 5. DỰ ĐOÁN MẪU MỚI
# ─────────────────────────────────────────────
def predict(model, texts: list) -> list:
    cleaned = [preprocess(t) for t in texts]
    preds = model.predict(cleaned)
    for text, label in zip(texts, preds):
        print(f"  [{label}] {text}")
    return preds.tolist()


# ─────────────────────────────────────────────
# 6. MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    df    = load_data(FILE_PATH)
    model = train(df)
    # model = joblib.load("model_vi_classification.pkl")  # Tải model đã lưu nếu đã train trước đó

    print("\n🔍 Thử dự đoán mẫu mới:")
    samples = [
        "thầy là nguyễn kim duy",   # ← thay bằng cụm từ thực tế
        "đọc tài liệu đi các em",
        "cô nè",
        "môn học này hay đó"
    ]
    predict(model, samples)