# K3s 部署指南 (Deploy Guide)

本教程將指導您如何將列印機查詢系統部署到現有的 K3s 集群 (Ubuntu 24.04 環境)。

## 📋 準備工作

確保您已經將最新的程式碼上傳到您的 K3s Manager 節點（或任意一台裝有 docker 和 k3s cli 的機器）。

目錄結構應包含：
- `Dockerfile`
- `requirements.txt`
- `webapp.py`
- `sharp_mfp_export.py`
- `k8s-deployment.yaml`
- `templates/`
- `static/`

---

## 🚀 步驟 1：構建 Docker 鏡像

在程式碼根目錄執行以下命令來構建鏡像：

```bash
# 構建鏡像，名稱為 printer-webapp:latest
sudo docker build -t printer-webapp:latest .
```

如果您沒有裝 Docker，可以使用 K3s 內建的 `ctr` 或 `nerdctl` (如果有的話)，但通常建議在開發生產機器上用 Docker 建置後匯出。

## 📦 步驟 2：將鏡像導入 K3s

因為 K3s 使用 containerd 作為容器運行時，它看不到本地 Docker 的鏡像。我們需要將鏡像保存並導入到 K3s 中。

**方法 A：在 K3s 節點上直接操作**

```bash
# 1. 將 Docker 鏡像保存為 tar 文件
sudo docker save printer-webapp:latest -o printer-webapp.tar

# 2. 將 tar 文件導入 K3s 的鏡像庫
sudo k3s ctr images import printer-webapp.tar
```

**確認導入成功：**
```bash
sudo k3s ctr images list | grep printer-webapp
# 應顯示 docker.io/library/printer-webapp:latest
```

## ☸️ 步驟 3：部署到 Kubernetes

我們已經準備好了 `k8s-deployment.yaml` 文件，其中包含了：
1.  **PVC**: 用於持久化存儲匯出的 Excel/CSV 數據。
2.  **Deployment**: 運行網站應用。
3.  **Service**: 開放端口讓外部訪問。

### 修改配置 (可選)
打開 `k8s-deployment.yaml`，檢查以下環境變數是否需要修改：
```yaml
        env:
        - name: SHARP_USER
          value: "admin"   <-- 修改為實際帳號
        - name: SHARP_PASS
          value: "admin"   <-- 修改為實際密碼
```

### 執行部署
```bash
sudo kubectl apply -f k8s-deployment.yaml
```

## ✅ 步驟 4：驗證與訪問

### 檢查 Pod 狀態
```bash
sudo kubectl get pods
# 等待 STATUS 變為 Running
```

### 檢查 Service
```bash
sudo kubectl get svc
```
您應該會看到 `printer-webapp-service`，類型為 `NodePort`，端口映射類似 `80:30080/TCP`。

### 訪問網站
現在，您可以通過 K3s 集群中 **任意節點的 IP** 加上端口 **30080** 來訪問系統。

例如：
`http://192.168.x.x:30080`

## 🔄 日後更新與維護

如果要更新程式碼：
1. 修改程式碼。
2. 重新構建鏡像：`sudo docker build -t printer-webapp:latest .`
3. 重新匯出導入：
   ```bash
   sudo docker save printer-webapp:latest -o printer-webapp.tar
   sudo k3s ctr images import printer-webapp.tar
   ```
4. 重啟 Pod 以應用新鏡像：
   ```bash
   sudo kubectl rollout restart deployment printer-webapp
   ```

## 🗑️ 故障排除

**查看日誌：**
```bash
# 先獲取 pod 名稱
sudo kubectl get pods
# 查看日誌
sudo kubectl logs printer-webapp-xxxxxxxxx-xxxxx
```

**進入容器內部：**
```bash
sudo kubectl exec -it printer-webapp-xxxxxxxxx-xxxxx -- /bin/bash
```
