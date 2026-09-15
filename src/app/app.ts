import { Component, inject, OnInit } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { FormsModule } from '@angular/forms';
import { SwUpdate } from '@angular/service-worker';
import { VERSION } from '../environments/version';


// import { SessionTimeoutService } from './services/session-timeout.service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, FormsModule],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App implements OnInit {

  private swUpdate = inject(SwUpdate);

  // private sessionTimeout = inject(SessionTimeoutService);

  ngOnInit(): void {
    console.log('app.ts -> App version:', VERSION.version);
    if (!this.swUpdate.isEnabled) {
      return;
    }
    this.swUpdate.checkForUpdate().then(updateFound => {
    console.log('Update check result:', updateFound);
  });

    this.swUpdate.versionUpdates.subscribe(event => {
      console.log('SW event:', event);
      switch (event.type) {

        case 'VERSION_DETECTED':
          console.log('New version detected');
          break;

        case 'VERSION_READY':
          if (confirm('🚀 A new version is available. Update now?')) {
            this.swUpdate.activateUpdate().then(() => {
              location.reload();
            });
          }
          break;

        case 'VERSION_INSTALLATION_FAILED':
          console.error('Update failed');
          break;
      }
    });

    // this.sessionTimeout.restoreSession();
  }
}