# HS Code Assistant — Kế hoạch triển khai chi tiết

> **Tên tệp:** `plant.md`  
> **Mục tiêu:** Xây dựng công cụ Web chạy Docker trên NAS dành cho 1–2 người dùng, đọc tờ khai XML/Excel đã thông quan, hình thành kho lịch sử hàng hóa, gợi ý HS Code, tạo mô tả khai báo và đưa ra căn cứ để người dùng xem xét.

---

## 0. Trạng thái khảo sát tệp nguồn

Đã phân tích trực tiếp tệp thật `ToKhaiHQ7X_QDTQ_308462191020.xls`.

- File là legacy `.xls` nhị phân BIFF8 trong OLE2 Compound Document, không phải `.xlsx`.
- Workbook có hai sheet: `TKX` hiển thị và `HANG` ở trạng thái very hidden.
- `TKX` có 7 trang in, 404 hàng, 31 cột và 9 dòng hàng.
- `HANG` là sheet mẫu bố cục trống cho hai dòng hàng, không phải dữ liệu thực.
- Dòng hàng được nhận diện bằng marker `<01>` đến `<09>`.
- Mã HS, mô tả, số lượng, đơn giá và trị giá có mapping tương đối ổn định trong từng item block.
- Model, nhãn hiệu, công dụng, tình trạng và xuất xứ chủ yếu nằm trong chuỗi mô tả, không có cột riêng.
- Số tiền và ngày tháng chủ yếu được lưu dạng chuỗi hiển thị; không có công thức Excel.
- Tổng 9 dòng hàng khớp tổng trị giá hóa đơn 54.000.000 VND.

Các đầu ra khảo sát:

```text
xls-structure-report.md
field-mapping-xls.csv
declaration-308462191020-structure.json
```

> **Chưa xác minh:** Mapping mới được xác nhận cho mẫu tờ khai xuất khẩu B11 này; chưa phải mapping chung cho tờ khai nhập khẩu, loại hình khác hoặc file từ phiên bản/phần mềm khác.

## 1. Phương án chốt

### 1.1. Hình thức sản phẩm

Ứng dụng chính là Web App nội bộ:

```text
Máy Windows của người dùng
        │ Trình duyệt Web
        ▼
NAS chạy Docker
├── Reverse Proxy
├── Frontend Web
├── Backend API
├── Worker xử lý nền
├── PostgreSQL + pgvector
├── Redis
├── Kho file trên NAS
└── AI Gateway
        │
        ▼
API AI bên ngoài — chỉ gọi khi người dùng yêu cầu
```

Không làm Desktop App đầy đủ trong MVP. Desktop Agent chỉ cân nhắc sau này nếu cần theo dõi thư mục XML của ECUS/phần mềm khai báo, tự tải file mới lên NAS, thông báo hoàn tất và mở nhanh trang chi tiết.

### 1.2. Công nghệ

| Thành phần | Công nghệ |
|---|---|
| Frontend | Next.js, TypeScript, React |
| UI | Tailwind CSS, shadcn/ui |
| Bảng dữ liệu | TanStack Table |
| Backend | Python, FastAPI, Pydantic, SQLAlchemy, Alembic |
| XML | `defusedxml` + `lxml` |
| Legacy XLS | `xlrd` hoặc BIFF reader tương đương, hỗ trợ merged cells và sheet visibility |
| XLSX | `openpyxl` |
| Database/search | PostgreSQL, pgvector, PostgreSQL FTS + trigram |
| Queue | Redis + Dramatiq hoặc Celery |
| Lưu file | Volume trực tiếp trên NAS |
| AI | API qua AI Gateway |
| Reverse proxy | Caddy hoặc Nginx Proxy Manager |
| Đóng gói | Docker Compose |

### 1.3. NAS tham khảo

> **Suy luận:** CPU x86-64 tối thiểu 4 nhân, RAM 16 GB khuyến nghị, SSD/NVMe 500 GB–1 TB cho Docker/PostgreSQL, HDD RAID1/SHR cho tài liệu/backup và nên có UPS. Không cần GPU, không chạy LLM local ở giai đoạn đầu.

## 2. Mục tiêu nghiệp vụ và giới hạn

```text
Nhập XML/Excel đã thông quan → bóc tách và chuẩn hóa
→ lưu tờ khai, dòng hàng và file gốc → nhận diện sản phẩm/model/part number
→ tìm lịch sử tương tự → tạo ứng viên HS → chấm điểm dữ liệu + rule
→ gọi AI theo yêu cầu → tạo mô tả, giải thích và căn cứ
→ người dùng duyệt/chỉnh sửa/khóa → xuất Excel
```

Mỗi dòng hàng cần có HS lịch sử, HS gợi ý chính, tối đa ba mã thay thế, điểm kỹ thuật, mức phù hợp, mô tả đề xuất, thuộc tính nhận diện, thông tin thiếu, tờ khai tương tự, căn cứ mở lại được, trạng thái duyệt và audit log.

Hệ thống chỉ hỗ trợ quyết định: không khẳng định kết luận pháp lý, không tự truyền tờ khai, không sửa dữ liệu khóa, không bịa căn cứ và luôn ghi rõ kết quả AI cần con người xem xét.

## 3. Phạm vi MVP

### Bắt buộc

- **Import:** nhiều XML/XLS/XLSX, template chuẩn và mapping Excel tự do; SHA-256 chống trùng; lưu file nguyên bản, parser/version; lỗi theo file/dòng.
- **Kho tờ khai:** danh sách/chi tiết/dòng hàng; tìm theo số tờ khai, invoice, ngày, loại hình; lọc HS, model, brand, xuất xứ.
- **Kho sản phẩm:** gom các dòng về hồ sơ chuẩn; liên kết model/part number/brand; lịch sử HS/mô tả; duyệt và khóa.
- **Gợi ý HS:** exact/fuzzy/full-text/vector search; ứng viên từ lịch sử; điểm minh bạch; cảnh báo xung đột.
- **AI thủ công:** phân tích, tạo mô tả, so sánh mã; JSON có schema; cache request hash; log token/trạng thái; không tự chạy cả file.
- **Export:** dòng chọn hoặc đã duyệt, mô tả, căn cứ, trạng thái kiểm tra và template tùy chỉnh.

### Ngoài MVP

OCR hàng loạt, tự đọc catalogue lớn, tích hợp trực tiếp VNACCS/ECUS DB, Windows Sync Agent, Mobile App, multi-tenant SaaS, LLM local và crawl nguồn không có API/cơ chế sử dụng ổn định.

## 4. Khảo sát và parser dữ liệu

### 4.1. Bộ mẫu

Cần 5–10 XML nhập và 5–10 XML xuất đã thông quan, nhiều loại hình, tờ khai một/nhiều dòng, cùng model khác mô tả, sản phẩm gần giống khác HS, tờ khai sửa đổi, Excel tương ứng và PDF/ảnh đối chiếu.

Dữ liệu thật không vào Git:

```text
sample-data-private/{raw,sanitized,expected,profiling}/
sample-data/{sanitized-xml,sanitized-excel,expected-json}/
```

### 4.2. XLS profiler và adapter `tkx_export_b11_xls_v1`

Phát hiện bằng magic bytes thay vì phần mở rộng:

```text
D0 CF 11 E0 A1 B1 1A E1  → OLE2/legacy XLS
PK 03 04                    → XLSX/ZIP
HTML/XML signature          → văn bản giả dạng XLS
```

Chọn sheet bằng anchor: cộng 5 điểm cho từng dấu hiệu “Tờ khai hàng hóa xuất khẩu”, “Số tờ khai”, “Mã số hàng hóa”, marker `<NN>`; trừ 10 nếu very hidden hoặc chỉ có nhãn mà không có HS thực. Kết quả chọn `TKX`, giữ `HANG` làm layout fingerprint.

Giả sử marker ở hàng `R`:

```text
HS Code                 F(R+2)
Mô tả                   F(R+3), merged F:AA trong 3 hàng
Số lượng (1)            Q(R+6)       Đơn vị (1) Y(R+6)
Số lượng (2)            Q(R+7)       Đơn vị (2) Y(R+7)
Trị giá hóa đơn         F(R+8)
Đơn giá hóa đơn         R(R+8)       Tiền tệ W(R+8)       Đơn vị Y(R+8)
Trị giá tính thuế       G(R+10)
Đơn giá tính thuế       R(R+11)      Tiền tệ Y(R+11)      Đơn vị AA(R+11)
```

Tìm marker bằng `^<\d{2}>$`, không dựa vào khoảng cách vì header lặp khi sang trang. Trước khi đọc offset phải xác nhận nhãn trong block.

Validation: số marker bằng tổng dòng; tổng trị giá dòng bằng invoice; `quantity_1 × invoice_unit_price = invoice_value`; sheet dữ liệu không very hidden; giữ raw string lẫn typed value; số tờ khai, MST và HS luôn là string.

### 4.3. XML profiler

Profiler phải đọc an toàn, phát hiện encoding/root/namespace, thống kê XPath và thuộc tính, nhóm lặp, kiểu mẫu đã ẩn, CDATA/chữ ký số, rồi xuất Markdown/JSON. Production cần giới hạn file, thời gian, depth/node; hash trước parse; không log dữ liệu nhạy cảm và báo lỗi encoding rõ ràng.

```python
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from defusedxml.lxml import fromstring
from lxml import etree


@dataclass
class XmlProfile:
    root_tag: str
    namespaces: dict[str, str]
    path_counts: dict[str, int]
    attribute_counts: dict[str, int]


def local_name(tag: str) -> str:
    return tag.split("}", 1)[1] if tag.startswith("{") else tag


def profile_xml(path: Path) -> XmlProfile:
    root = fromstring(path.read_bytes())
    paths: Counter[str] = Counter()
    attributes: Counter[str] = Counter()

    def walk(node: etree._Element, parent: str) -> None:
        current = f"{parent}/{local_name(node.tag)}"
        paths[current] += 1
        for key in node.attrib:
            attributes[f"{current}/@{local_name(key)}"] += 1
        for child in node:
            if isinstance(child.tag, str):
                walk(child, current)

    walk(root, "")
    namespaces = {prefix or "default": uri for prefix, uri in (root.nsmap or {}).items() if uri}
    return XmlProfile(local_name(root.tag), namespaces, dict(paths), dict(attributes))
```

### 4.4. Parser adapter và regression

Không tạo một parser với hàng chục nhánh nguồn. Dùng `BaseDeclarationParser` cùng adapter theo source/version và một `GenericMappingParser`. Mỗi adapter khai báo ID/version, fingerprint, namespace, XPath mapping, parse header/items, validation và warnings.

```python
from typing import Protocol


class DeclarationParser(Protocol):
    parser_id: str
    parser_version: str

    def can_parse(self, document: bytes) -> bool: ...
    def parse(self, document: bytes) -> "NormalizedDeclaration": ...
```

Mỗi XML sanitized có expected JSON. Test số tờ khai/dòng/HS, bảo toàn mô tả, Decimal, namespace, trùng file, external entity và lỗi XML có kiểm soát.

## 5. Mô hình chuẩn hóa

Luôn lưu dữ liệu gốc, dữ liệu chuẩn hóa, parser version, dữ liệu AI và dữ liệu người dùng duyệt; AI không ghi đè dữ liệu gốc.

```python
from datetime import date, datetime
from decimal import Decimal
from pydantic import BaseModel, Field


class NormalizedGoodsItem(BaseModel):
    line_number: int
    hs_code_original: str | None = None
    description_original: str
    product_name: str | None = None
    english_name: str | None = None
    model: str | None = None
    part_number: str | None = None
    brand: str | None = None
    manufacturer: str | None = None
    material: str | None = None
    function: str | None = None
    specifications: dict[str, str] = Field(default_factory=dict)
    origin_code: str | None = None
    quantity: Decimal | None = None
    unit_code: str | None = None
    unit_price: Decimal | None = None
    customs_value: Decimal | None = None
    currency_code: str | None = None
    tax_rate: Decimal | None = None
    raw_fields: dict[str, object] = Field(default_factory=dict)


class NormalizedDeclaration(BaseModel):
    declaration_number: str
    declaration_date: datetime | None = None
    declaration_type: str | None = None
    direction: str | None = None
    customs_office_code: str | None = None
    clearance_status: str | None = None
    importer_tax_code: str | None = None
    exporter_name: str | None = None
    invoice_number: str | None = None
    invoice_date: date | None = None
    parser_id: str
    parser_version: str
    items: list[NormalizedGoodsItem]
    warnings: list[str] = Field(default_factory=list)
    raw_metadata: dict[str, object] = Field(default_factory=dict)
```

Trạng thái thông quan có `source`, `raw`, `normalized`, `verified_by_user`, `verified_at`. Nếu nguồn không đủ rõ thì để `unknown` và yêu cầu xác nhận.

## 6. Thiết kế database

- `source_files`: metadata, storage path, MIME, size, SHA-256 unique, source type, parser/version, status/error, uploader/time.
- `declarations`: file nguồn, số/ngày/loại/direction, cửa khẩu, trạng thái và nguồn trạng thái, dữ liệu bên liên quan đã mask, invoice, currency/value, parser/version, raw JSONB. Unique dự kiến `(declaration_number, declaration_date, direction)`.
- `declaration_items`: line unique trong declaration; HS/mô tả gốc; text/thuộc tính chuẩn hóa; Decimal dưới `NUMERIC`; specifications/raw JSONB.
- `products`: hồ sơ chuẩn, approved HS/description, review status, lock metadata và timestamps.
- `product_item_links`: khóa ghép product/item, method/score và thông tin người xác nhận.
- `hs_codes`: code/version, mô tả VI/EN, hierarchy, thời hạn hiệu lực và source document.
- `hs_suggestions`: candidate/rank/score components/sources/rules/status/reviewer.
- `ai_requests`: request hash/type/entity/provider/model/prompt version; redacted input/output; token/cost/status/error/audit.
- `evidence_sources`, `evidence_chunks`: nguồn, hiệu lực, URL/hash, vị trí/page/section/content/embedding.
- `audit_logs`: user/action/entity, old/new JSONB, IP, user agent và timestamp.

Extensions/index chính:

```sql
CREATE EXTENSION IF NOT EXISTS pg_trgm;
CREATE EXTENSION IF NOT EXISTS vector;
CREATE INDEX idx_items_model_trgm
ON declaration_items USING gin (model gin_trgm_ops);
```

## 7. Chuẩn hóa và tìm kiếm hàng hóa

Không thay `description_original`. Pipeline: Unicode normalization → khoảng trắng/ký tự → model/part-number rule → brand dictionary → specs → normalized search text → người dùng kiểm tra. Giữ cả original/normalized và không xóa sớm `/`, `-`, `.`, đơn vị hay mã phiên bản.

Ưu tiên model: trường riêng → part number riêng → từ khóa `model`, `mã`, `P/N`, `part no.` → regex → AI nếu rule không đủ. Không coi mọi chuỗi chữ-số là model.

Thứ tự search: exact part number → exact model + brand → exact normalized model → trigram → full-text → vector → AI rerank thủ công. Embedding chỉ gồm thuộc tính phân loại; không gồm số tờ khai, doanh nghiệp, invoice hay trị giá. Trọng số hybrid phải hiệu chỉnh bằng golden dataset.

## 8. Bộ máy gợi ý HS và Rule Engine

Không yêu cầu LLM chọn trực tiếp từ toàn biểu HS. Tìm lịch sử/product trước, tổng hợp mã đã dùng, lấy top ứng viên, áp rule, chấm điểm rồi chỉ nhờ AI giải thích/so sánh.

Nguồn ứng viên: cùng part number; model + brand; product master; hàng tương tự đã thông quan; danh mục nội bộ; văn bản/quyết định trong kho; mã người dùng thêm. Lưu chi tiết từng thành phần điểm và phạt vì xung đột, thiếu thuộc tính phân loại, mô tả chung/ngắn hoặc lịch sử cũ/không rõ trạng thái.

Mức hiển thị: **Cao theo dữ liệu nội bộ**, **Khá phù hợp**, **Cần kiểm tra**, **Thiếu dữ liệu**, **Có xung đột**; không dùng “chắc chắn đúng”.

Rule cấu hình có version, điều kiện, required fields, candidate headings và warnings:

```yaml
rule_id: POWER_CONVERTER_001
version: 1
name: Thiết bị chuyển đổi nguồn AC/DC
match:
  any:
    - field: function
      operator: contains
      value: chuyển đổi AC sang DC
    - field: normalized_text
      operator: regex
      value: "(?i)AC\\s*/?\\s*DC|power supply"
required_fields: [input_voltage, output_voltage, rated_power]
candidate_headings: ["8504"]
warnings:
  - Chưa xác định có chức năng sạc pin hay không
```

Kết quả lưu `rule_set_version`, rules applied/passed/failed và warnings. Rule mới không tự sửa kết quả đã duyệt mà chỉ đề xuất tái đánh giá.

## 9. AI Gateway

- Frontend không giữ key; backend chỉ gửi payload sản phẩm tối thiểu đã redact.
- Structured output bắt buộc; không ghi đè kết quả đã duyệt; provider cấu hình được.
- Các tác vụ: chuẩn hóa mô tả, tạo mô tả khai báo, so sánh 3–5 ứng viên, tóm tắt căn cứ.
- Không tự thêm model, công suất, vật liệu, tình trạng, xuất xứ hay công dụng thiếu nguồn.
- Mọi claim căn cứ phải có `source_id`.
- Request hash gồm payload, candidate codes, evidence IDs/hashes, prompt version, provider và model.
- Cache vô hiệu khi input/evidence/prompt/model đổi hoặc người dùng yêu cầu.
- Mặc định loại tên doanh nghiệp, MST, số tờ khai, invoice, trị giá và đối tác.
- Chế độ: `AI_DISABLED`, `AI_MANUAL_ONLY`, `AI_ASSISTED`; MVP dùng `AI_MANUAL_ONLY`.

Output gồm `recommended_code`, trạng thái cần review, mô tả, thuộc tính, evidence có source ID, alternative codes kèm điều kiện, missing information, warnings và `human_review_required: true`.

## 10. API

```text
POST /api/auth/login                 POST /api/auth/logout
GET  /api/auth/me                    POST /api/auth/change-password

POST /api/imports/xml                POST /api/imports/excel
GET  /api/imports                    GET  /api/imports/{job_id}
GET  /api/imports/{job_id}/errors    POST /api/excel-mappings

GET  /api/declarations               GET  /api/declarations/{id}
GET  /api/declarations/{id}/items    POST /api/declarations/{id}/verify-clearance

GET/POST /api/products               GET/PATCH /api/products/{id}
POST /api/products/{id}/merge        POST /api/products/{id}/approve
POST /api/products/{id}/lock         POST /api/products/{id}/unlock

POST /api/products/{id}/find-similar GET /api/products/{id}/similar-items
POST /api/products/{id}/suggest-hs   GET /api/products/{id}/hs-suggestions
POST /api/hs-suggestions/{id}/approve|reject

POST /api/ai/normalize-product       POST /api/ai/generate-description
POST /api/ai/compare-codes           GET /api/ai/requests
GET  /api/ai/usage

POST/GET /api/evidence               GET /api/evidence/{id}
POST /api/evidence/{id}/index        DELETE /api/evidence/{id}

POST /api/exports/excel              GET /api/exports/{id}
GET  /api/exports/{id}/download
```

## 11. Giao diện

- **Dashboard:** tổng tờ khai/dòng/sản phẩm; xung đột, thiếu dữ liệu, chờ duyệt; AI requests và import gần đây.
- **Import:** XML, Excel, lịch sử, mapping; kết quả thành công/trùng/lỗi, counts, warning và parser.
- **Tờ khai:** số, ngày, loại hình, thông quan, invoice, số dòng, file và parser version.
- **Hàng hóa:** tên/model/part/brand, HS lịch sử/gợi ý, mức phù hợp, xung đột, trạng thái, thao tác.
- **Chi tiết sản phẩm:** chuẩn hóa, HS + mô tả, lịch sử, căn cứ, AI, audit; nút tìm, gợi ý, AI, tạo mô tả, duyệt và khóa.
- **So sánh HS:** lịch sử, model match, similarity, rule, evidence, thiếu dữ liệu và AI explanation.
- **Cấu hình:** provider/model/key status, limit/timeout, template, backup, parser/rule versions.

## 12. Docker trên NAS

Containers: `hs-proxy`, `hs-frontend`, `hs-backend`, `hs-worker`, `hs-postgres`, `hs-redis`; chưa cần MinIO.

```text
/volume1/docker/hs-assistant/
├── compose/{docker-compose.yml,.env,secrets/}
├── postgres/                     ├── redis/
├── uploads/{incoming,originals,sanitized,quarantine}/
├── evidence/                     ├── exports/
├── backups/                      ├── logs/
└── temp/
```

Compose production dùng pgvector/pg16, Redis append-only, health checks, internal network, persistent bind mounts và secrets. Bổ sung image pin, resource limits, log rotation, backup job, AI secret, security headers, read-only filesystem nơi phù hợp và non-root user.

> **Suy luận cho RAM 16 GB:** PostgreSQL 2–3 GB, worker 2–3 GB, backend 1–2 GB, frontend 0,5–1 GB, Redis 256–512 MB, proxy dưới 256 MB; theo dõi thực tế trước khi đặt hard limits.

## 13. Bảo mật

- **XML:** chặn DTD/external entity/network, giới hạn size/node/depth, không mặc định `huge_tree=True`, escape khi render, quarantine file lỗi.
- **Excel:** giới hạn size, không macro/công thức, cảnh báo `.xlsm`, sanitize tên file/sheet; chống formula injection với giá trị bắt đầu `=`, `+`, `-`, `@` khi export.
- **AI:** secret phía server, payload preview đã ẩn, request ID không secret, timeout/retry hữu hạn và circuit breaker.
- **App:** Argon2id; cookie HttpOnly/Secure/SameSite; login rate limit; CSRF cho cookie session; Admin/User RBAC; audit; không expose PostgreSQL/Redis.
- **Remote:** ưu tiên WireGuard/Tailscale/VPN NAS, không port-forward database.

## 14. Backup, logging và vận hành

Backup PostgreSQL, file gốc, evidence, cấu hình, rules, mappings, templates và secrets đã mã hóa/quy trình tái tạo. Khởi điểm: dump + snapshot hằng ngày, bản sao thiết bị khác hằng tuần, restore test định kỳ. RAID không phải backup; cần bản chính, bản khác thiết bị và bản off-site/cloud mã hóa.

Log upload/hash/parser/warning/import, AI, suggestion, approve/reject/lock, export và login fail; không log key, password, XML nguyên bản hay dữ liệu đối tác không cần thiết.

Health endpoints: `/health/live`, `/health/ready`, `/health/database`, `/health/redis`. Theo dõi CPU/RAM, disk, DB size, queue, import/AI failures và backup status; có thể dùng Uptime Kuma.

## 15. Cấu trúc source

```text
hs-code-assistant/
├── frontend/{app,components,features,lib,types}/
├── backend/
│   ├── app/{api,auth,core,database,models,schemas,imports,parsers,
│   │        normalization,matching,hs_engine,rules,ai,evidence,exports,audit,workers}/
│   ├── alembic/
│   └── tests/
├── deployment/{docker-compose.yml,Caddyfile,secrets,scripts,backup}/
├── rules/              ├── mappings/          ├── templates/
├── sample-data/        ├── docs/              └── .github/workflows/
```

## 16. Kế hoạch theo cổng nghiệm thu

1. **Gate 0 — Khảo sát:** XML/namespace profile, mapping, schema variants, sanitized sample, expected JSON. Xác định items, đối chiếu counts/fields và cách nhận biết thông quan.
2. **Gate 1 — Docker:** Compose, migration, login, storage, health. Restart không mất dữ liệu, DB không expose, restore backup thành công.
3. **Gate 2 — Parser:** adapters, Excel mapping, jobs/errors. Tất cả mẫu khớp expected JSON, chống trùng/XXE và xử lý namespace.
4. **Gate 3 — Kho lịch sử:** declaration UI, product master, linking/search. Tìm model/part, xem mọi lần khai và xác nhận link.
5. **Gate 4 — HS không AI:** candidates, hybrid search, score components, conflicts. Không tạo mã ngoài pool; mọi điểm có nguồn; exact match ưu tiên.
6. **Gate 5 — AI thủ công:** Gateway, schema, cache, token/cost, redaction. Chỉ chạy khi bấm, không lộ key, reject sai schema, lỗi AI không mất dữ liệu.
7. **Gate 6 — Duyệt/xuất:** approval, lock, audit, Excel. Không ghi đè dữ liệu khóa, đúng template và truy được người sửa/duyệt.

## 17. Bộ kiểm thử

- **Unit:** XML paths/namespaces, Decimal/date, HS/model normalization, hashing, redaction, scoring.
- **Integration:** upload → queue → parser → DB; Excel mapping; search → candidates; mocked AI/schema; approve/lock; export.
- **Security:** XXE, XML bomb/oversize, ZIP bomb, formula injection, path traversal, unauthorized download, key exposure, brute-force login.
- **Golden dataset:** input product, expected history/candidates/warnings/questions và human-approved code. Không đánh giá bằng cảm giác “AI nghe hợp lý”.

## 18. Kết quả và quyết định sau khảo sát

Đã xác minh BIFF8/OLE2, `TKX` là dữ liệu, `HANG` very hidden là template, report-layout dùng anchor/marker, HS 8 chữ số, mô tả merged, model/brand không riêng, xuất xứ ở hậu tố `#&XX`, số liệu chủ yếu là chuỗi, trạng thái ở tiêu đề, 7 trang/9 dòng/tổng 54.000.000 VND và không công thức.

Chưa xác minh: nhập khẩu, loại hình khác B11, sửa đổi, generator khác, trên 99 dòng, nhiều tiền tệ/thuế và XML tương ứng.

Quyết định:

- Legacy TKX dùng `tkx_export_b11_xls_v1`, magic bytes + sheet + anchors, block offset có validation; `HANG` chỉ là fingerprint.
- XML dùng profiler rồi adapter theo schema/version; đối chiếu XML/XLS để chọn nguồn ưu tiên từng trường.
- Excel bảng phẳng dùng Generic Column Mapping Parser, không dùng adapter TKX.

## 19. Checklist đầu vào

- [ ] Một XML nhập khẩu đã thông quan.
- [ ] Một XML xuất khẩu đã thông quan.
- [ ] Một file nhiều dòng hàng.
- [x] Một XLS B11 xuất khẩu đã thông quan.
- [x] Generator: GemBox.Spreadsheet 4.1 for .NET.
- [ ] Xác nhận được phép phân tích nội bộ.
- [ ] Xác nhận trường cần ẩn khi đưa sample vào Git.

Đã có `xls-structure-report.md`, `field-mapping-xls.csv`, `declaration-308462191020-structure.json`. Đang chờ XML để tạo `xml-profile.json`, `xml-structure-report.md`, `field-mapping-xml.csv`, sanitized sample, expected declaration, adapter và tests.

## 20. Kết luận

Phương án phù hợp nhất là Web chạy Docker Compose trên NAS, FastAPI, PostgreSQL + pgvector, Redis worker, file volume trực tiếp và AI API manual-only. Chất lượng phụ thuộc trước hết vào parser đúng dữ liệu thật, lịch sử chuẩn hóa, liên kết đúng sản phẩm/model/part number, điểm có thể giải thích, căn cứ thật, quy trình duyệt/khóa và nguyên tắc AI không bổ sung dữ liệu thiếu nguồn.
