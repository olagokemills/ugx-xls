import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { exportJsonToExcel, exportJsonToCsv } from 'ugx-xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  template: `
    <main class="container">
      <h1>ugx-xlsx Demo</h1>
      <p>Click a button to download the report.</p>
      <div class="actions">
        <button (click)="exportExcel()">Export to Excel</button>
        <button (click)="exportCsv()">Export to CSV</button>
      </div>

      <h3>Preview</h3>
      <pre>{{ mapForExport() | json }}</pre>
    </main>
  `,
  styles: [
    `
      .container {
        padding: 24px;
        font-family: system-ui, Arial, sans-serif;
      }
      .actions {
        display: flex;
        gap: 12px;
        margin: 12px 0 24px;
      }
      button {
        padding: 8px 12px;
        cursor: pointer;
      }
      pre {
        background: #f6f8fa;
        padding: 12px;
        border-radius: 6px;
      }
    `,
  ],
})
export class AppComponent {
  private FCYReports = [
    {
      RequestDate: '2025-10-26',
      ValueDate: '2025-10-27',
      SubsidiaryName: 'Lagos HQ',
      AccountId: '123-4567890',
      Initiator: 'Jane Doe',
      Amount: 350000,
      Charges: 200,
      AccountNumber: '0123456789',
      BeneficiaryName: 'John Smith',
      BankName: 'First Bank',
      Narration: 'Payment for goods',
      ApprovalStatus: 'Approved',
      PaymentStatus: 'Completed',
      PaymentRemark: 'OK',
    },
  ];

  mapForExport() {
    return this.FCYReports.map((res) => ({
      'Request Date': res.RequestDate,
      'Value Date': res.ValueDate,
      'Subsidiary Name': res.SubsidiaryName,
      'Source Account': res.AccountId,
      Initiator: res.Initiator,
      Amount: res.Amount,
      Charge: res.Charges,
      'Beneficiary Account Number': res.AccountNumber,
      'Beneficiary Name': res.BeneficiaryName,
      'Beneficiary Bank': res.BankName,
      Narration: res.Narration,
      'Approval Status': res.ApprovalStatus,
      'Payment Status': res.PaymentStatus,
      'Payment Remark': res.PaymentRemark,
    }));
  }

  exportExcel(): void {
    exportJsonToExcel(
      this.mapForExport(),
      `Single-Reports-${new Date().toISOString().slice(0, 10)}`
    );
  }

  exportCsv(): void {
    exportJsonToCsv(
      this.mapForExport(),
      `Single-Reports-${new Date().toISOString().slice(0, 10)}`,
      { includeBom: true }
    );
  }
}
